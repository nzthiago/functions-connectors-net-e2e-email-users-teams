using Azure.Core;
using Azure.Identity;
using Azure.Monitor.OpenTelemetry.Exporter;
using Company.Function;
using Azure.Connectors.Sdk.Office365;
using Azure.Connectors.Sdk.Office365Users;
using Azure.Connectors.Sdk.Teams;
using Microsoft.Azure.Functions.Worker.OpenTelemetry;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;


var applicationInsightsConnectionString = Environment.GetEnvironmentVariable("APPLICATIONINSIGHTS_CONNECTION_STRING");

// DefaultAzureCredential works in both environments:
//   - In Azure: uses the user-assigned managed identity whose client id is in AZURE_CLIENT_ID.
//   - Locally:  falls back to the developer's `az login` / VS / VS Code credentials.
// (Plain ManagedIdentityCredential would try IMDS at 169.254.169.254 — fine in Azure, fails locally.)
var credential = new DefaultAzureCredential(new DefaultAzureCredentialOptions
{
    ManagedIdentityClientId = Environment.GetEnvironmentVariable("AZURE_CLIENT_ID")
});

var host = new HostBuilder()
    .ConfigureFunctionsWebApplication()
    .ConfigureServices(services =>
    {
        // OpenTelemetry → Azure Monitor (App Insights) for the worker process.
        // Paired with `"telemetryMode": "OpenTelemetry"` in host.json so host + worker
        // telemetry stays correlated. 
        var openTelemetry = services.AddOpenTelemetry()
            .UseFunctionsWorkerDefaults();

        if (!string.IsNullOrWhiteSpace(applicationInsightsConnectionString))
        {
            openTelemetry.UseAzureMonitorExporter(o =>
            {
                o.ConnectionString = applicationInsightsConnectionString;
                o.Credential = credential;
            });
        }

        services.AddSingleton<TokenCredential>(credential);

        services.AddSingleton(sp => new TeamsClient(
            new Uri(RequireSetting("TEAMS_CONNECTION_RUNTIME_URL")),
            sp.GetRequiredService<TokenCredential>()));

        // Office 365 client — used both for sender-history enrichment (GetEmailsAsync)
        // and to flag the source email (FlagAsync) once we decide it's important.
        // Same connection runtime URL the trigger uses; just consumed as a client too.
        services.AddSingleton(sp => new Office365Client(
            new Uri(RequireSetting("OFFICE365_CONNECTION_RUNTIME_URL")),
            sp.GetRequiredService<TokenCredential>()));

        // Office 365 Users client — used to look up the sender's M365 profile
        // (UserProfileAsync + ManagerAsync) for IN-ORG badging and card enrichment.
        services.AddSingleton(sp => new Office365UsersClient(
            new Uri(RequireSetting("OFFICE365USERS_CONNECTION_RUNTIME_URL")),
            sp.GetRequiredService<TokenCredential>()));

        services.AddSingleton<ImportanceClassifier>();
    })
    .Build();

host.Run();

// Fail loudly at boot if a required connection URL setting is missing — otherwise
// the connector clients silently get an empty BaseAddress and every call throws
// the cryptic "An invalid request URI was provided" deep in the request pipeline.
static string RequireSetting(string name)
{
    var value = Environment.GetEnvironmentVariable(name);
    if (string.IsNullOrWhiteSpace(value))
    {
        throw new InvalidOperationException(
            $"Required app setting '{name}' is not set. " +
            $"For local development add it to function-app/local.settings.json. " +
            $"For Azure deployments it should be wired via infra/main.bicep.");
    }
    return value;
}

