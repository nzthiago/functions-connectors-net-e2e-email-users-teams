# Demo guide — Important Email Triage with Connector Namespace

A two-part walkthrough:

1. **[Part 1 — Local debug](#part-1--local-debug-with-vs-code)** — show the code paths and decisions step-by-step in the VS Code debugger using `func start` + a synthetic payload from [test.http](../test.http).
2. **[Part 2 — Cloud walkthrough](#part-2--cloud-walkthrough)** — show the deployed resources in the Azure Portal, the Connector Namespace configuration, and trigger an end-to-end run from a real email.

> Architecture refresher: ![architecture](architecture.png)
> Source: [docs/architecture.drawio](architecture.drawio)

---

## Part 1 — Local debug with VS Code

### Prerequisites

- `azd up` already run and connections authorized (see [README](../README.md)).
- [Azurite](https://learn.microsoft.com/azure/storage/common/storage-use-azurite) running (the VS Code extension's "Start Azurite" command works great).
- `az login --tenant <your tenant>` (so `DefaultAzureCredential` can reach the Connector Namespace runtime URLs locally).
- A populated [function-app/local.settings.json](../function-app/local.settings.json) — values come from `azd env get-values` plus the `TEAMS_CONNECTION_RUNTIME_URL`, `OFFICE365_CONNECTION_RUNTIME_URL`, and `MSGRAPHGROUPSANDUSER_CONNECTION_URL` app settings on the deployed Function App.
- The C# Dev Kit + Azure Functions VS Code extensions.

### Launch the host under the debugger

1. Open the workspace in VS Code.
2. **Run and Debug** → "Attach to .NET Functions" (or just hit **F5** from the function-app folder; the Azure Functions extension scaffolds a launch config the first time).
3. Wait for the function listing to print:
   ```
   Functions:
     OnNewImportantEmailReceived: [POST] http://localhost:7071/api/OnNewImportantEmailReceived
   ```

### Set the test.http variables for local

At the top of [test.http](../test.http):

```
@functionAppUrl = http://localhost:7071
@httpPostCode =
```

### The breakpoint tour

Set these breakpoints in order, then click **Send Request** above the **first** POST in [test.http](../test.http) (the urgent one).

#### 🔴 BP 1 — Trigger entry: see the raw connector payload
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var emails = payload.Body?.Value ?? [];`

- Inspect `payload` in **Locals**. This is the exact JSON the Connector Namespace POSTs back, deserialized by the **source-generated** `Office365OnNewEmailTriggerPayload` from the Connectors SDK — no manual JSON parsing.
- `Body.Value` is a list — the gateway can batch multiple new emails per webhook call.
- Hover `payload.Body.Value[0]` to show the strongly-typed `GraphClientReceiveMessage`.

#### 🔴 BP 2 — The classifier decision
**[function-app/ImportanceClassifier.cs](../function-app/ImportanceClassifier.cs)** — top of `Classify()`, `var reasons = new List<string>();`

- Step through (**F10**) the three checks one by one and watch `reasons` grow.
- Hover `_importantSenders` — `HashSet<string>` populated once at startup from the `IMPORTANT_SENDERS` env var (case-insensitive lookup).
- At the `ScoreContent` call, step **into** (**F11**) — that's where the regex weights add up.

#### 🔴 BP 3 — Regex matching loop
**[function-app/ImportanceClassifier.cs](../function-app/ImportanceClassifier.cs)** — `foreach (var (pattern, weight, label) in UrgencyPatterns)`

- Watch `score` and `matchedLabels` grow. The first test payload (`"Urgent: action needed on connector demo"`) should hit "urgency adverb" and "explicit action/decision required".
- **Conditional breakpoint trick:** right-click → **Edit Breakpoint** → condition `pattern.IsMatch(haystack)` to stop only on hits.
- Talking point: adding a new signal is a one-line tuple in `UrgencyPatterns` — easy to extend.

#### 🔴 BP 4 — The skip path
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `if (!verdict.IsImportant)`

- After the urgent run, hit **Send Request** on the **second** POST in [test.http](../test.http) (newsletter). Verdict has zero reasons → 200 OK with no Teams call. Cheap negative path; keeps Teams quiet.

#### 🔴 BP 5 — Sender history via Office 365
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var response = await _office365Client.GetEmailsAsync(...)` (inside `FetchFromFolderAsync`)

- This is a **live call** to the Office 365 connector — proves managed-identity / `DefaultAzureCredential` is working locally.
- Notice it's the **same connector** the trigger uses, just consumed as a client. One authorized connection, two roles: trigger source + outbound API.
- Step out to `GetSenderHistoryAsync` and show the `Task.WhenAll` over `["Inbox", "Archive"]` — `GetEmailsAsync` is per-folder by design, so we fan out two parallel calls and merge the results. Easy to extend to `SentItems` etc. by adding to the array.
- Step over and inspect `response.Value` — list of `GraphClientReceiveMessage`. The `from: senderEmail` arg is a **server-side filter** on the watched mailbox, so it scales to any tenant size — there's no directory enumeration involved.
- Hover `history` after construction — show `TotalRecent` / `LastWeek` / `MostRecent`, which feed straight into the Teams card.

#### 🔴 BP 6 — In-team/external badge via directory group lookup
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — inside `IsSenderInWatchedGroupAsync(...)`

- Directory-driven: `WATCHED_GROUP_ID` points to a Microsoft Entra security group object ID, and the Microsoft Graph Groups & Users connector checks direct members with `mail eq '<sender>'`.
- Results are cached per cold start for 10 minutes, so bursty mail from the same sender doesn't repeatedly hit the connector. Group member → `🟢 IN-TEAM`; non-member → `🔴 EXTERNAL`; no `WATCHED_GROUP_ID` → no badge.

#### 🔴 BP 7 — Teams card composition
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var request = new PostMessageRequest { ... }`

- Show the `PostMessageRequest : DynamicPostMessageRequest` derived class. The SDK uses an empty marker base + your subclass to serialize the operation's dynamic schema. **Zero hand-rolled JSON.**
- Inspect `messageBody` to show the assembled HTML with badge, role, groups, and "Why flagged".

#### 🔴 BP 8 — The actual outbound call
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — `var result = await _teamsClient.PostMessageToConversationAsync(...)`

- One line of code → an authenticated REST call to the Teams connection runtime URL via the gateway.
- Step over and switch to your Teams channel — the card appears in real time.
- Inspect `result.MessageID` — strongly-typed response.

#### 🔴 BP 9 — Write-back: flag the source email
**[function-app/ProcessEmail.cs](../function-app/ProcessEmail.cs)** — inside `FlagSourceMessageAsync`, `await _office365Client.FlagAsync(...)`

- This is the "connectors do things, not just read" moment. Same Office 365 connection, this time mutating mailbox state — sets the Outlook follow-up flag on the source message.
- Step over and switch to Outlook — the original email now has a red flag, so even if the recipient misses the Teams ping they have a server-side reminder.
- Failures are caught + logged at Warning — a flag failure never breaks the triage flow.

### Bonus: catch connector failures fast

**Debug → Windows → Exception Settings** → tick `Office365ConnectorException` and `TeamsConnectorException`. If a connection isn't authorized, you'll stop right at the failure with `ex.StatusCode` visible.

### Bonus talking points to weave in

- **DI in [Program.cs](../function-app/Program.cs)** — `TeamsClient`, `Office365Client`, and `ImportanceClassifier` are all singletons; the connector clients take `(connectionRuntimeUrl, TokenCredential)`.
- **One Office 365 connection, three uses** — the trigger callback, the sender-history `GetEmailsAsync` lookup, and the `FlagAsync` write-back all flow through the same authorized connection. No extra OAuth dance per capability.
- **No `Microsoft.Graph` SDK** — mailbox history still uses mailbox-scoped Office 365 calls, while the `msgraphgroupsanduser` connector provides an authoritative direct group-membership check for the in-team badge.
- **Deterministic local repro** — the typed trigger payload means [test.http](../test.http) is enough to exercise the full pipeline; no real Outlook needed during development. (The flag write-back will fail with a fake MessageId in test.http — that's expected and just logs a warning.)

---

## Part 2 — Cloud walkthrough

Now switch over to the deployed app to show how the same code runs in Azure and gets called by a real Office 365 email.

### Where things live

```
Subscription
└── rg-importantemailprocessor
    ├── func-<token>                  Function App (Flex Consumption)
    ├── plan-<token>                  Function App Plan (FC1)
    ├── id-<token>                    User-Assigned Managed Identity
    ├── st<token>                     Storage account (deployment + AzureWebJobsStorage)
    ├── log-<token>                   Log Analytics workspace
    ├── appi-<token>                  Application Insights
    └── cns-<token>                   Connector Namespace  (lives in brazilsouth)
        ├── connections/
        │   ├── cnsc-<token>             Office 365 Outlook (trigger + client)
        │   ├── cnsc-teams-<token>       Microsoft Teams
        │   └── cnsc-msgraph-<token>     Microsoft Graph Groups & Users
        └── triggerconfigs/
            └── cnsc-<token>-trigger    Office 365 OnNewEmailV3 → callback URL
```

Run `azd env get-values` to grab the live names if you need them.

### A) Tour of the Function App

[Azure Portal → Resource Group → Function App].

1. **Overview** — point out it's **Flex Consumption** (`FC1`, 2 GB memory, scale-to-zero) and the system-assigned + user-assigned MI.
2. **Settings → Environment variables** — show the relevant ones:
   - `TEAMS_TEAM_ID`, `TEAMS_CHANNEL_ID`, `IMPORTANT_SENDERS`, `WATCHED_GROUP_ID` — same values you saw in `local.settings.json`.
   - `TEAMS_CONNECTION_RUNTIME_URL`, `OFFICE365_CONNECTION_RUNTIME_URL`, `MSGRAPHGROUPSANDUSER_CONNECTION_URL` — the gateway endpoints the SDK clients hit.
   - `AZURE_CLIENT_ID` — the UAMI client id; the connector clients use it to acquire tokens.
   - `APPLICATIONINSIGHTS_AUTHENTICATION_STRING: ClientId=...;Authorization=AAD` — AAD-only telemetry, no instrumentation key in plaintext.
3. **Settings → Identity → User assigned** — show the same UAMI; click into it to show its role assignments (Storage Blob Data Owner, Queue/Table Data Contributor on the storage account, Monitoring Metrics Publisher on App Insights).
4. **Functions → OnNewImportantEmailReceived → Code + Test → Logs** — leave it streaming for the real-email step below.
5. **Functions → App Keys → System keys** — show `connector_extension`. The post-deploy script grabbed this value to build the trigger callback URL — that's how the gateway is allowed to invoke the function.

### B) Tour of the Connector Namespace

[Azure Portal → Resource Group → cns-`<token>`].

1. **Overview** — show it's a brand-new resource type (`Microsoft.Web/connectorNamespaces@2026-05-01-preview`) and that it lives in `brazilsouth` while the function lives in `westus2`. Cross-region by design today.
2. **Connections** — three rows:
   - `cnsc-<token>` (Office 365)
   - `cnsc-teams-<token>` (Teams)
   - `cnsc-msgraph-<token>` (Microsoft Graph Groups & Users)

   Click each → **Status: Connected** (you authorized them after `azd up`). Show the **Access policies** tab — the Function App's UAMI principal id is allowed to use the connection. That's how the SDK clients get to call the runtime URL with no shared secret. The Office 365 connection has access policies because the function calls `GetEmailsAsync` and `FlagAsync` against it directly, in addition to receiving the trigger callback. The Graph Groups & Users connection requires Microsoft Entra ID consent for the group-membership lookup.
3. **Trigger configurations** → `cnsc-<token>-trigger`:
   - **Operation:** `OnNewEmailV3`
   - **Parameters:** `folderPath = Inbox` (note: no `importance` filter — we evaluate every mail in code via the classifier so we can use richer signals).
   - **Notification details / callback URL:** `https://func-<token>.azurewebsites.net/runtime/webhooks/connector?functionName=OnNewImportantEmailReceived&code=<system key>`. This URL was assembled by [infra/scripts/postdeploy.sh](../infra/scripts/postdeploy.sh) using the system key from step A.5.

### C) Tour of Application Insights

[Azure Portal → Resource Group → appi-`<token>`].

1. **Live metrics** — leave it open in a side panel for the real-email demo.
2. **Logs (KQL)** — handy queries to have ready:

   ```kusto
   // last 1h of important emails accepted
   traces
   | where timestamp > ago(1h)
   | where message has "Important email accepted"
   | project timestamp, message, severityLevel
   | order by timestamp desc
   ```

   ```kusto
   // skip-vs-accept ratio
   traces
   | where timestamp > ago(24h)
   | where message has_any ("Important email accepted", "Skipping non-important email")
   | summarize count() by tostring(split(message, ".")[0])
   ```

   ```kusto
   // any connector failures
   exceptions
   | where timestamp > ago(24h)
   | where type endswith "ConnectorException"
   | project timestamp, type, outerMessage, customDimensions
   ```

### D) Trigger a real run

1. **Pick a sender from the allowlist** (or temporarily add yourself) — confirm with `az functionapp config appsettings list -g rg-importantemailprocessor -n func-<token> --query "[?name=='IMPORTANT_SENDERS'].value" -o tsv`.
2. **Send yourself an email** from Outlook (web or desktop) to the inbox the Office 365 connection was authorized for. Two good test cases:
   - **Allowlist hit:** subject `Quick check on connector demo` from one of the `IMPORTANT_SENDERS`. Will trigger reason "Sender in IMPORTANT_SENDERS allowlist".
   - **Content heuristic hit:** any sender, subject `[URGENT] please review by EOD` — triggers "bracketed urgency tag" + "deadline language".
   - **Skip case:** any non-allowlisted sender, plain subject like `lunch tomorrow?` — should NOT post to Teams.
3. **What to watch happen, in order:**
   1. **Application Insights → Live metrics**: a new Request appears within ~5–30 s of the email landing (Office 365 trigger polling cadence).
   2. **Function App → Functions → OnNewImportantEmailReceived → Logs**: you'll see
      - `Important email accepted. Subject=... From=... Reasons=...` (for hits), or
      - `Skipping non-important email. Subject=... From=... Importance=...` (for misses).
      - For accepted emails, also `Flagged source email. MessageId=...`.
   3. **Teams channel**: the triage card pops in with the optional IN-TEAM/EXTERNAL badge, sender history, "Why flagged", subject, and preview.
   4. **Outlook**: the source email now shows the red follow-up flag.
4. **Show the end-to-end correlation**: in App Insights → Transaction search, click the request → see the dependency calls out to `*.logic-df.azure-apihub.net` (the Office 365 `GetEmailsAsync` + `FlagAsync` calls and the Teams `PostMessageToConversationAsync` call), with timing.

### E) Talking points for the cloud part

- **No secrets anywhere.** Function → connections is via UAMI access policy; function → storage / App Insights via UAMI role assignments; connection → external service via the user who authorized it in the gateway. Nothing in app settings is sensitive.
- **Server-side filter is just `folderPath = Inbox`.** All the importance logic is in your code — easy to evolve without redeploying gateway config.
- **Same code, two execution modes.** The local debug session and the cloud function run **identical** code paths and call the **same** connection runtime URLs. Local just uses your user identity instead of the UAMI.
- **Diagnostics surface naturally.** Every accept/skip is a structured log line, and connector failures throw typed exceptions you can pivot on in App Insights.

### F) Reset / cleanup

- Remove a sender from the allowlist:
  ```bash
  azd env set IMPORTANT_SENDERS "<comma list without that address>"
  azd deploy   # or azd provision if you only want to update settings
  ```
- Tear everything down:
  ```bash
  azd down --purge
  ```
