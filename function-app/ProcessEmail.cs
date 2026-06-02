using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.Functions.Worker.Extensions.Connector;
using Azure.Connectors.Sdk.Office365;
using Azure.Connectors.Sdk.Office365Users;
using Azure.Connectors.Sdk.Teams;
using System.Collections.Concurrent;
using System.Text.Json;
using Azure.Connectors.Sdk.Office365.Models;
using Azure.Connectors.Sdk.Teams.Models;
using Azure.Connectors.Sdk;

namespace Company.Function
{
    /// <summary>
    /// Triage bot for inbound mail. The Office 365 trigger fires for every new message
    /// in the watched folder; this function decides — via <see cref="ImportanceClassifier"/>
    /// — whether each one is important enough to surface in the Teams triage channel.
    ///
    /// Important means ANY of:
    ///   * Sender is in the IMPORTANT_SENDERS allowlist
    ///   * Email is flagged High importance by the sender
    ///   * Subject/body contains urgency / action-required language above threshold
    ///
    /// For mails that pass the bar we also use the Office 365 connector — the same
    /// connection that fires the trigger — to:
    ///   * Pull the sender's recent message history from the watched mailbox so the
    ///     Teams card can show "N emails in last 7d, last seen X" context.
    ///   * Flag the source email server-side so the user has a follow-up reminder
    ///     in Outlook even if they miss the Teams notification.
    /// Sender org/external badging uses an `INTERNAL_DOMAINS` (comma-separated) prefilter
    /// to keep the Office 365 Users API off the hot path for clearly external mail. When
    /// the sender's domain matches one in INTERNAL_DOMAINS, we call UserProfileAsync to
    /// confirm + enrich the card with job title, department, and manager. When the domain
    /// doesn't match (or the API returns 404 even for a matching domain), the sender is
    /// treated as EXTERNAL. If INTERNAL_DOMAINS is unset, every sender is looked up.
    /// </summary>
    public class ProcessEmail
    {
        private const string PostAsFlowBot = "Flow bot";
        private const string PostInChannel = "Channel";
        private const int SenderHistoryDays = 7;
        private const int SenderHistoryFetchTop = 25;
        private const int SenderProfileCacheTtlMinutes = 10;

        // Enrichment data fetched from the Office 365 Users connector.
        private sealed record SenderProfile(
            string? DisplayName,
            string? JobTitle,
            string? Department,
            string? ManagerDisplayName);

        // Maps normalised sender email → (profile, notFound, cachedAt).
        //   profile != null            → in-org sender, enrichment available
        //   profile == null, notFound  → external sender (404 from UserProfileAsync)
        //   profile == null, !notFound → transient lookup failure, badge omitted
        private static readonly ConcurrentDictionary<string, (SenderProfile? profile, bool notFound, DateTime cachedAt)> SenderProfileCache = new(StringComparer.OrdinalIgnoreCase);

        private sealed record SenderHistory(int TotalRecent, int LastWeek, DateTime? MostRecent)
        {
            public static SenderHistory Empty { get; } = new(0, 0, null);
        }

        private readonly ILogger _logger;
        private readonly TeamsClient _teamsClient;
        private readonly Office365Client _office365Client;
        private readonly Office365UsersClient _office365UsersClient;
        private readonly ImportanceClassifier _classifier;
        private readonly string _teamsTeamId;
        private readonly string _teamsChannelId;
        private readonly IReadOnlyList<string> _internalDomains;

        public ProcessEmail(
            ILoggerFactory loggerFactory,
            TeamsClient teamsClient,
            Office365Client office365Client,
            Office365UsersClient office365UsersClient,
            ImportanceClassifier classifier)
        {
            _logger = loggerFactory.CreateLogger<ProcessEmail>();
            _teamsClient = teamsClient;
            _office365Client = office365Client;
            _office365UsersClient = office365UsersClient;
            _classifier = classifier;
            _teamsTeamId = Environment.GetEnvironmentVariable("TEAMS_TEAM_ID") ?? "";
            _teamsChannelId = Environment.GetEnvironmentVariable("TEAMS_CHANNEL_ID") ?? "";
            _internalDomains = (Environment.GetEnvironmentVariable("INTERNAL_DOMAINS") ?? "")
                .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(d => d.ToLowerInvariant())
                .ToArray();
        }

        // True when no INTERNAL_DOMAINS allowlist is configured (look up every sender) or when
        // the sender's domain matches the allowlist. False means we should skip the API call
        // and treat the sender as external.
        private bool ShouldLookupSender(string senderEmail)
        {
            if (_internalDomains.Count == 0) return true;
            var atIdx = senderEmail.LastIndexOf('@');
            if (atIdx < 0 || atIdx == senderEmail.Length - 1) return false;
            var domain = senderEmail[(atIdx + 1)..].Trim().ToLowerInvariant();
            return _internalDomains.Any(d => domain == d || domain.EndsWith("." + d));
        }

        [Function("OnNewImportantEmailReceived")]
        public async Task<IActionResult> OnNewImportantEmailReceived(
            [ConnectorTrigger()] Office365OnNewEmailTriggerPayload payload)
        {
            var emails = payload?.Body?.Value;
            _logger.LogInformation(
                "Trigger callback received. emailCount={Count}",
                emails?.Count ?? -1);

            if (emails is null || emails.Count == 0)
            {
                _logger.LogWarning("Empty trigger payload — nothing to process.");
                return new OkResult();
            }

            foreach (var email in emails)
            {
                if (email is null) continue;

                var verdict = _classifier.Classify(
                    email.From,
                    email.Subject,
                    email.Body,
                    email.BodyPreview,
                    email.Importance);

                if (!verdict.IsImportant)
                {
                    _logger.LogInformation(
                        "Skipping non-important email. Subject={Subject} From={From} Importance={Importance}",
                        email.Subject, email.From, email.Importance);
                    continue;
                }

                _logger.LogInformation(
                    "Important email accepted. Subject={Subject} From={From} Reasons={Reasons}",
                    email.Subject, email.From, string.Join(" | ", verdict.Reasons));

                var history = await GetSenderHistoryAsync(email.From);
                await PostTriageCardAsync(email, history, verdict);
                await FlagSourceMessageAsync(email);
            }

            return new OkResult();
        }

        /// <summary>
        /// Pulls the sender's recent history from the watched mailbox via
        /// <see cref="Office365Client.GetEmailsAsync"/>, fanning out across Inbox
        /// and Archive (the connector's GetEmails is per-folder). Scales to any
        /// tenant size — it's a server-side filter on `from` against the connected
        /// mailbox, not a directory call. Best-effort: failures degrade gracefully
        /// to "no history".
        /// </summary>
        private async Task<SenderHistory> GetSenderHistoryAsync(string? senderEmail)
        {
            if (string.IsNullOrWhiteSpace(senderEmail))
            {
                return SenderHistory.Empty;
            }

            string[] folders = ["Inbox", "Archive"];
            var perFolderTasks = folders.Select(f => FetchFromFolderAsync(senderEmail, f));
            var perFolderResults = await Task.WhenAll(perFolderTasks);

            var messages = perFolderResults.SelectMany(r => r).ToList();
            if (messages.Count == 0)
            {
                return SenderHistory.Empty;
            }

            var cutoff = DateTime.UtcNow.AddDays(-SenderHistoryDays);
            var lastWeek = messages.Count(m => m.ReceivedTime is DateTime t && t >= cutoff);
            var mostRecent = messages
                .Select(m => m.ReceivedTime)
                .Where(t => t.HasValue)
                .DefaultIfEmpty()
                .Max();

            return new SenderHistory(messages.Count, lastWeek, mostRecent);
        }

        private async Task<IReadOnlyList<GraphClientReceiveMessage>> FetchFromFolderAsync(string senderEmail, string folder)
        {
            try
            {
                var response = await _office365Client.GetEmailsAsync(
                    folder: folder,
                    to: null,
                    cC: null,
                    toOrCC: null,
                    from: senderEmail,
                    importance: null,
                    onlyWithAttachments: false,
                    subjectFilter: null,
                    fetchOnlyUnreadMessages: false,
                    originalMailboxAddress: null,
                    includeAttachments: false,
                    searchQuery: null,
                    top: SenderHistoryFetchTop,
                    cancellationToken: default);

                return (IReadOnlyList<GraphClientReceiveMessage>?)response?.Value ?? [];
            }
            catch (ConnectorException ex)
            {
                _logger.LogWarning(ex,
                    "Office365 GetEmails failed for sender {Sender} in folder {Folder}. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage} — skipping that folder.",
                    senderEmail, folder, ex.ConnectorName, ex.ErrorCode, ex.Message);
                return [];
            }
        }

        /// <summary>
        /// Sets the Outlook follow-up flag on the source message via the Office 365
        /// connector. Gives the recipient a server-side reminder regardless of whether
        /// they see the Teams card. Best-effort — flag failures don't fail the function.
        /// </summary>
        private async Task FlagSourceMessageAsync(GraphClientReceiveMessage email)
        {
            if (string.IsNullOrEmpty(email.MessageId))
            {
                _logger.LogDebug("No MessageId on payload; skipping flag.");
                return;
            }

            try
            {
                await _office365Client.FlagAsync(
                    messageId: email.MessageId,
                    input: new UpdateEmailFlag { Flag = new { flagStatus = "flagged" } },
                    originalMailboxAddress: null,
                    cancellationToken: default);

                _logger.LogInformation("Flagged source email. MessageId={MessageId}", email.MessageId);
            }
            catch (ConnectorException ex)
            {
                _logger.LogWarning(ex,
                    "Failed to flag source email. MessageId={MessageId}. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage}",
                    email.MessageId, ex.ConnectorName, ex.ErrorCode, ex.Message);
            }
        }

        /// <summary>
        /// Looks up the sender's M365 user profile via the Office 365 Users connector.
        /// A successful lookup means the sender is in the org — we get rich enrichment for
        /// free (job title, department, manager). A 404 / <see cref="ConnectorException"/> with StatusCode 404
        /// means the sender is external. Other exceptions degrade gracefully: badge is omitted.
        /// Results are cached for 10 minutes to absorb bursty mail volume.
        /// </summary>
        private async Task<(SenderProfile? profile, bool notFound)> GetSenderProfileAsync(string? senderEmail)
        {
            if (string.IsNullOrWhiteSpace(senderEmail))
                return (null, false);

            var normalizedSender = senderEmail.Trim();
            var now = DateTime.UtcNow;

            if (SenderProfileCache.TryGetValue(normalizedSender, out var cached) &&
                now - cached.cachedAt < TimeSpan.FromMinutes(SenderProfileCacheTtlMinutes))
            {
                return (cached.profile, cached.notFound);
            }

            // Domain prefilter: if INTERNAL_DOMAINS is configured and the sender's domain
            // is not in it, skip the API call and treat the sender as external.
            if (!ShouldLookupSender(normalizedSender))
            {
                SenderProfileCache[normalizedSender] = (null, true, now);
                return (null, true);
            }

            try
            {
                var user = await _office365UsersClient.UserProfileAsync(normalizedSender);

                // Fetch manager display name — best-effort, users with no manager return null or throw.
                string? managerName = null;
                try
                {
                    var manager = await _office365UsersClient.ManagerAsync(normalizedSender);
                    managerName = manager?.DisplayName;
                }
                catch (ConnectorException)
                {
                    // No manager record — tolerated.
                }

                var profile = new SenderProfile(user?.DisplayName, user?.JobTitle, user?.Department, managerName);
                SenderProfileCache[normalizedSender] = (profile, false, now);
                return (profile, false);
            }
            catch (ConnectorException ex)
            {
                _logger.LogWarning(ex,
                    "Office 365 Users profile lookup failed for sender {Sender}. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage}", 
                    normalizedSender, ex.ConnectorName, ex.ErrorCode, ex.Message);
                return (null, false);
            }
        }

        private async Task PostTriageCardAsync(GraphClientReceiveMessage email, SenderHistory history, ImportanceVerdict verdict)
        {
            if (string.IsNullOrEmpty(_teamsTeamId) || string.IsNullOrEmpty(_teamsChannelId))
            {
                _logger.LogWarning("TEAMS_TEAM_ID or TEAMS_CHANNEL_ID not configured. Skipping Teams notification.");
                return;
            }

            var (senderProfile, senderNotFound) = await GetSenderProfileAsync(email.From);

            var badge = senderNotFound
                ? "🔴 <b>EXTERNAL — verify identity before acting</b><br/>"
                : senderProfile is not null
                    ? "🟢 <b>IN-ORG</b><br/>"
                    : "";

            var profileLine = senderProfile is not null
                ? $"<br/><b>Title:</b> {senderProfile.JobTitle ?? "(not set)"}" +
                  $" | <b>Dept:</b> {senderProfile.Department ?? "(not set)"}" +
                  (senderProfile.ManagerDisplayName is not null
                      ? $" | <b>Manager:</b> {senderProfile.ManagerDisplayName}"
                      : "")
                : "";

            var historyLine = history.TotalRecent switch
            {
                0 => "<br/><b>Sender history:</b> no prior emails from this sender in Inbox or Archive",
                _ => $"<br/><b>Sender history:</b> {history.TotalRecent} emails from this sender across Inbox + Archive " +
                     $"({history.LastWeek} in last {SenderHistoryDays}d" +
                     (history.MostRecent is DateTime t ? $", most recent {t:yyyy-MM-dd HH:mm} UTC" : "") +
                     ")"
            };

            var reasonsLine = verdict.Reasons.Count > 0
                ? $"<br/><b>Why flagged:</b> {string.Join("; ", verdict.Reasons)}"
                : "";

            var messageBody =
                $"<b>📧 Email triage — review required</b><br/>" +
                $"{badge}" +
                $"<b>From:</b> {email.From}{profileLine}{historyLine}{reasonsLine}<br/>" +
                $"<b>Subject:</b> {email.Subject}<br/>" +
                $"<b>Preview:</b> {email.BodyPreview ?? "(no preview)"}<br/>" +
                $"<i>(source email has been flagged in Outlook)</i>";

            // DynamicPostMessageRequest is a *dynamic-schema* body whose properties
            // are resolved at runtime by the connector's schema discovery endpoint.
            // The connector backend expects camelCase keys (`recipient.groupId`,
            // `recipient.channelId`, `messageBody`). System.Text.Json preserves the
            // declared casing of strongly-typed properties on subclasses, so the
            // safe way to populate the body is via AdditionalProperties (decorated
            // with [JsonExtensionData] on the base class) using literal camelCase
            // keys — matching the official sample at
            // https://github.com/Azure/Connectors-NET-Samples/blob/main/DirectConnector/TeamsFunctions.cs
            var request = new DynamicPostMessageRequest();
            request.AdditionalProperties["recipient"] = JsonSerializer.SerializeToElement(
                new
                {
                    groupId = _teamsTeamId,
                    channelId = _teamsChannelId,
                });
            request.AdditionalProperties["messageBody"] = JsonSerializer.SerializeToElement(messageBody);

            try
            {
                var result = await _teamsClient.PostMessageToConversationAsync(
                    PostAsFlowBot,
                    PostInChannel,
                    request);

                _logger.LogInformation("Triage card posted to Teams. MessageId={MessageId}", result?.MessageId);
            }
            catch (ConnectorException ex)
            {
                _logger.LogError(ex, "Failed to post Teams message. ConnectorName={ConnectorName}, ErrorCode={ErrorCode}, ErrorMessage={ErrorMessage}",  ex.ConnectorName, ex.ErrorCode, ex.Message);
            }
        }
    }
}
