# Azure Functions with Connector Namespace - Email Processor sample with Users and Teams actions

This sample demonstrates how to use **Azure Functions** with **Connector Namespace connectors** to react to events from external services. It listens for new emails arriving in a Microsoft 365 inbox, classifies each one with a small in-process importance heuristic, and — for the ones that pass the bar — enriches the message with **sender history** from the same mailbox, posts a formatted card to a **Microsoft Teams** channel, and **flags the source email** in Outlook so the recipient also has a server-side follow-up reminder.

## Architecture

![Architecture diagram](docs/architecture.png)

> Editable source: [docs/architecture.drawio](docs/architecture.drawio) (open with [draw.io](https://app.diagrams.net)).

- **Azure Functions (Flex Consumption)** — A .NET 10 isolated worker function app that receives HTTP callbacks from the Connector Namespace.
- **Connector Namespace** — Manages three connections (Office 365, Teams, Office 365 Users) and the Office 365 trigger configuration.
- **Office 365 Outlook Connector** — Used in two ways:
  - As a **trigger** — the gateway watches the Inbox (`folderPath: Inbox`) and calls the function for every new email.
  - As a **client** inside the function — `GetEmailsAsync` to fetch sender history (last N messages from the same sender), and `FlagAsync` to set the Outlook follow-up flag on the source email when it's classified as important. Scales to any tenant size because everything is scoped to the watched mailbox — no directory enumeration required.
- **Office 365 Users Connector** — Looks up the sender's M365 user profile (`UserProfileAsync`) to determine whether they are in the org. A successful lookup means the sender is an org user (🟢 IN-ORG badge); a 404 means external (🔴 EXTERNAL badge). When the profile is found the card is also enriched with the sender's job title, department, and manager display name (via `ManagerAsync`). To keep this API off the hot path for clearly external mail, configure the optional `INTERNAL_DOMAINS` setting (comma-separated, e.g. `microsoft.com,contoso.com`) — only senders whose domain matches will be looked up. Leave it empty to look up every sender.
- **Teams Connector** — Posts the enriched triage card to a configured Teams channel.

## Prerequisites

- [Azure Developer CLI (azd)](https://learn.microsoft.com/azure/developer/azure-developer-cli/install-azd)
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli)
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local)
- [jq](https://jqlang.github.io/jq/) (required by the post-deploy script on Linux/macOS)
- An Azure subscription
- A Microsoft 365 account (for the email, users, and Teams connectors)

## Getting Started

### 1. Clone This Repository

```bash
git clone <url>/FunctionAppConnectorsEmailProcessor
cd FunctionAppConnectorsEmailProcessor
```

### 2. Deploy to Azure

Before running `azd up`, set the Teams team and channel IDs as `azd` environment variables so the deployment can wire them into the Function App settings (`TEAMS_TEAM_ID` / `TEAMS_CHANNEL_ID`).

**Get the Team ID and Channel ID from Microsoft Teams:**

1. Open Microsoft Teams (desktop or web).
2. Find the channel where you want triage cards posted.
3. Click the channel's **⋯ (More options)** menu → **Get link to channel** → **Copy**.
4. The copied URL looks like:

   ```
   https://teams.microsoft.com/l/channel/19%3aXXXXXXXXXXXXXXXX%40thread.tacv2/General?groupId=00000000-1111-2222-3333-444444444444&tenantId=...
   ```

   - **Team ID** → the value of the `groupId` query parameter (e.g. `00000000-1111-2222-3333-444444444444`).
   - **Channel ID** → the segment after `/channel/`, URL-decoded (replace `%3a` with `:` and `%40` with `@`), e.g. `19:XXXXXXXXXXXXXXXX@thread.tacv2`.

**Set the azd env vars:**

```bash
azd env set TEAMS_TEAM_ID "<your-team-id>"
azd env set TEAMS_CHANNEL_ID "<your-channel-id>"
```

**Deploy:**

```bash
azd up
```

This provisions all infrastructure (Function App, Connector Namespace, Storage, Application Insights) and deploys the function code. After deployment, a post-deploy script automatically creates the Connector Namespace trigger configuration.

### 3. Authorize the Connections

The post-deploy script already kicks off connection authorization via the Azure CLI (the portal authorization UX is not yet available for Connector Namespace connections). For each of the three connections — **Office 365**, **Teams**, and **Office 365 Users** — the script:

1. Installs the early preview [`connector-namespace` Azure CLI extension](https://github.com/anthonychu/azure-cli-extensions/releases/tag/connector-namespace-0.1.0) if it isn't already present.
2. Runs `az connector-namespace connection authorize`, which opens a browser tab for OAuth consent.

Sign in with the appropriate account at each prompt:

- **Office 365** — the account whose inbox you want to monitor (drives the trigger, sender-history lookup, and follow-up flag).
- **Teams** — an account that can post to the target Teams channel.
- **Office 365 Users** — an account that can read user profiles in the directory (used for `UserProfileAsync` / `ManagerAsync`).

If you ever need to re-run authorization manually (e.g. token expired, account changed), use the commands below. The `*_CONNECTION_NAME` values are exported by the deployment outputs and stored in your `azd` env (`azd env get-values`):

```bash
# Office 365 (trigger + sender history + follow-up flag)
az connector-namespace connection authorize \
  --resource-group "$AZURE_RESOURCE_GROUP" \
  --namespace-name "$connectorNamespaceName" \
  --name "$connectorNamespaceConnectionName"

# Teams (post triage card)
az connector-namespace connection authorize \
  --resource-group "$AZURE_RESOURCE_GROUP" \
  --namespace-name "$connectorNamespaceName" \
  --name "$connectorNamespaceTeamsConnectionName"

# Office 365 Users (IN-ORG badge + manager enrichment)
az connector-namespace connection authorize \
  --resource-group "$AZURE_RESOURCE_GROUP" \
  --namespace-name "$connectorNamespaceName" \
  --name "$connectorNamespaceOffice365usersConnectionName"
```

> The `connector-namespace` CLI extension is installed automatically by the post-deploy script. To install it manually:
>
> ```bash
> az extension add \
>   --source https://github.com/anthonychu/azure-cli-extensions/releases/download/connector-namespace-0.1.0/connector_namespace-0.1.0-py2.py3-none-any.whl \
>   --yes
> ```

Until each connection is authorized, the trigger will not fire, Teams notifications will fail, and/or the IN-ORG badge will be omitted.

## Environment Variables

| Name | Description |
|---|---|
| `TEAMS_TEAM_ID` | Teams team/group ID where triage cards are posted. |
| `TEAMS_CHANNEL_ID` | Teams channel ID where triage cards are posted. |
| `IMPORTANT_SENDERS` | Optional comma-separated email allowlist whose messages always count as important. |

### 4. Test the Solution

Once the connections are authorized, send an email to the authorized account. The function classifies it via [function-app/ImportanceClassifier.cs](function-app/ImportanceClassifier.cs); for important ones it (1) calls the Office 365 connector to get the sender's recent history across the Inbox and Archive folders, (2) posts an enriched triage card to the configured Teams channel, and (3) flags the source email in Outlook.

You can also manually test the function endpoint using the [test.http](test.http) file (update the URL and function key to match your deployment).

## Project Structure

| Path | Description |
|---|---|
| `function-app/` | Azure Functions application (.NET 10, isolated worker) |
| `function-app/ProcessEmail.cs` | Function triggered for every new email; classifies importance, looks up sender history via the Office 365 connector, posts to Teams, and flags the source email |
| `function-app/Program.cs` | Host builder, registers Teams, Office 365, and Office 365 Users connector clients |
| `infra/main.bicep` | Main Bicep template for all Azure resources |
| `infra/connectorNamespace.bicep` | Connector Namespace plus Office 365, Teams, and Office 365 Users connection resources |
| `infra/scripts/postdeploy.sh` | Post-deploy script (Linux/macOS) — creates the Office 365 trigger config |
| `infra/scripts/postdeploy.ps1` | Post-deploy script (Windows) — creates the Office 365 trigger config |
| `azure.yaml` | Azure Developer CLI project configuration |
| `test.http` | Sample HTTP request for manual testing |

## Resources

- [Azure Functions documentation](https://learn.microsoft.com/azure/azure-functions/)
- [Azure Functions Flex Consumption plan](https://learn.microsoft.com/azure/azure-functions/flex-consumption-plan)
- [Azure Developer CLI (azd)](https://learn.microsoft.com/azure/developer/azure-developer-cli/)
