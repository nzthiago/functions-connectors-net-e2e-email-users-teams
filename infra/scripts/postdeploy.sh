#!/bin/bash
# Post-deployment configuration for the Functions + Connector Namespace sample.
#
# Uses the official `connector-namespace` Azure CLI extension from
# https://github.com/Azure/Connectors. Two responsibilities:
#   1. Create the Office 365 OnNewEmailV3 trigger config that POSTs new
#      emails to the function's connector webhook URL.
#   2. Walk the operator through OAuth consent for each of the three
#      connections (Office 365 Outlook, Teams, Office 365 Users) by
#      opening the consent link in a browser and polling until the
#      connection flips to `Connected`.
#
# Connection access policies for the function-app MI and the deployer
# user are created by Bicep (infra/connectorNamespace.bicep), so this
# script does not grant ACLs.

set -e

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
WHITE='\033[1;37m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Post-deployment configuration...${NC}"

# --- Read azd outputs --------------------------------------------------------
outputs=$(azd env get-values --output json)

if ! command -v jq &> /dev/null; then
    echo -e "${RED}Error: jq is required for this script. Please install jq.${NC}"
    exit 1
fi

subscriptionId=$(echo "$outputs" | jq -r '.AZURE_SUBSCRIPTION_ID')
resourceGroupName=$(echo "$outputs" | jq -r '.resourceGroupName')
connectorNamespaceName=$(echo "$outputs" | jq -r '.connectorNamespaceName')
connectorNamespaceConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceConnectionName')
connectorNamespaceTeamsConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceTeamsConnectionName')
connectorNamespaceOffice365usersConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceOffice365usersConnectionName')
functionAppName=$(echo "$outputs" | jq -r '.functionAppName')
office365FunctionName=$(echo "$outputs" | jq -r '.office365FunctionName')

# --- Install the official connector-namespace az CLI extension --------------
# Resolve the latest released wheel URL from the Azure/Connectors GitHub
# releases. All releases are marked pre-release so the standard
# /releases/latest endpoint 404s; we fetch /releases?per_page=1 instead.
# Pin a specific version by setting CONNECTOR_NAMESPACE_EXT_URL in the
# environment before running azd up.
if [[ -z "${CONNECTOR_NAMESPACE_EXT_URL:-}" ]]; then
    CONNECTOR_NAMESPACE_EXT_URL=$(curl -fsSL \
        "https://api.github.com/repos/Azure/Connectors/releases?per_page=1" \
        | grep -oE '"browser_download_url"\s*:\s*"[^"]*connector_namespace[^"]*\.whl"' \
        | head -1 \
        | sed 's/.*"\(https[^"]*\)".*/\1/')
fi
if [[ -z "${CONNECTOR_NAMESPACE_EXT_URL:-}" ]]; then
    echo -e "${RED}ERROR: could not resolve connector-namespace extension URL from Azure/Connectors releases${NC}" >&2
    exit 2
fi
if [[ -z "$(az extension show --name connector-namespace --query name -o tsv 2>/dev/null || true)" ]]; then
    echo -e "${CYAN}Installing 'connector-namespace' Azure CLI extension from $CONNECTOR_NAMESPACE_EXT_URL${NC}"
    az extension add --upgrade --yes --source "$CONNECTOR_NAMESPACE_EXT_URL"
fi

# --- Create Connector Namespace trigger config ------------------------------
echo ""
echo -e "${YELLOW}Creating Connector Namespace trigger config...${NC}"

echo -e "${CYAN}Fetching connector extension key for ${functionAppName}...${NC}"
# The system key is generated when the Functions runtime loads the
# Microsoft.Azure.Functions.Worker.Extensions.Connectors extension,
# which can take a few seconds after the function app finishes
# deploying. Poll for it briefly before giving up.
connectorExtensionKey=""
deadline=$(($(date +%s) + 180))
while [[ $(date +%s) -lt $deadline ]]; do
    connectorExtensionKey=$(az functionapp keys list -g "${resourceGroupName}" -n "${functionAppName}" --query "systemKeys.connector_extension" -o tsv 2>/dev/null || echo "")
    if [[ -n "${connectorExtensionKey}" ]]; then break; fi
    echo -e "${YELLOW}  connector_extension key not yet present; waiting 10s...${NC}"
    sleep 10
done
if [[ -z "${connectorExtensionKey}" ]]; then
    echo -e "${RED}ERROR: connector_extension system key never appeared on ${functionAppName}. The Microsoft.Azure.Functions.Worker.Extensions.Connectors extension may have failed to load. Check the function app's log stream.${NC}" >&2
    exit 1
fi

triggerName="${connectorNamespaceConnectionName}-trigger"
callbackUrl="https://${functionAppName}.azurewebsites.net/runtime/webhooks/connector?functionName=${office365FunctionName}&code=${connectorExtensionKey}"

connectionDetails=$(jq -nc --arg conn "${connectorNamespaceConnectionName}" \
    '{connectorName:"office365", connectionName:$conn}')
parameters=$(jq -nc '[{name:"folderPath", value:"Inbox"}]')
# The connector-namespace extension's --notification-details parser
# expects `body` to be a dict (not the Logic Apps template string
# `@triggerBody()`), so we omit it and let the runtime use the
# trigger's default output as the POST body. Auth + callback URL
# go in via the notification-details JSON.
#
# We write the dict-shaped args to temp files and pass them with the
# `@file` syntax. Passing them inline as JSON strings makes the CLI's
# shorthand parser try to interpret the leading `{` as the shorthand
# `key=value` syntax, which fails on colons inside URLs.
notificationDetails=$(jq -nc --arg url "${callbackUrl}" '{callbackUrl:$url, httpMethod:"Post"}')
connDetailsFile=$(mktemp)
notifDetailsFile=$(mktemp)
echo "${connectionDetails}"   > "${connDetailsFile}"
echo "${notificationDetails}" > "${notifDetailsFile}"
trap 'rm -f "${connDetailsFile}" "${notifDetailsFile}"' EXIT

echo -e "${CYAN}  Trigger name: ${triggerName}${NC}"
echo -e "${CYAN}  Callback URL: ${callbackUrl}${NC}"

# Best-effort idempotency: delete any prior config with this name.
az connector-namespace trigger delete \
    -g "${resourceGroupName}" \
    --namespace "${connectorNamespaceName}" \
    -n "${triggerName}" \
    --yes 2>/dev/null || true

az connector-namespace trigger create \
    -g "${resourceGroupName}" \
    --namespace "${connectorNamespaceName}" \
    -n "${triggerName}" \
    --connection-details "@${connDetailsFile}" \
    --operation-name "OnNewEmailV3" \
    --parameters "${parameters}" \
    --notification-details "@${notifDetailsFile}" \
    --state "Enabled" \
    --description "When a new email arrives in the consented Office 365 mailbox, POST the payload to the function's connector webhook."

echo -e "${GREEN}✅ Connector Namespace trigger config created successfully!${NC}"

# --- Authorize the connector connections (OAuth consent) --------------------
# Portal authorization UX is not yet available for Connector Namespace
# connections, so we drive OAuth consent through the CLI:
#   1. `connection list-consent-links` returns a one-shot logic-apis
#      consent URL with `state` already baked in.
#   2. We open the URL in a browser. The user signs in, consent is
#      persisted server-side, and the page redirects to portal.azure.com
#      (the well-known target the consent service has special-cased
#      for "no app to redirect to" CLI flows).
#   3. Poll `connection show --query properties.overallStatus` until it
#      flips to `Connected`.
echo ""
echo -e "${YELLOW}Authorizing connector connections via Azure CLI...${NC}"

open_url() {
    local url="$1"
    if command -v xdg-open >/dev/null 2>&1; then
        xdg-open "$url" >/dev/null 2>&1 || true
    elif command -v open >/dev/null 2>&1; then
        open "$url" >/dev/null 2>&1 || true
    elif command -v wslview >/dev/null 2>&1; then
        wslview "$url" >/dev/null 2>&1 || true
    fi
}

authorize_connection() {
    local connectionName="$1"
    local description="$2"

    echo -e "${CYAN}-> Authorizing ${description} connection: ${connectionName}${NC}"

    local currentStatus
    currentStatus=$(az connector-namespace connection show \
        -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
        -n "${connectionName}" --query "properties.overallStatus" -o tsv 2>/dev/null || echo "")
    if [[ "$(echo "$currentStatus" | tr '[:upper:]' '[:lower:]')" == "connected" ]]; then
        echo -e "${GREEN}   already Connected; skipping consent flow${NC}"
        return
    fi

    local params consentJson link paramsFile
    # `list-consent-links` --parameters: the URL value contains a colon
    # which makes the CLI's shorthand parser barf. Write to a temp file
    # and pass with the @file syntax.
    paramsFile=$(mktemp)
    echo '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]' > "${paramsFile}"
    consentJson=$(az connector-namespace connection list-consent-links \
        -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
        --connection-name "${connectionName}" --parameters "@${paramsFile}" -o json 2>/dev/null || echo "")
    rm -f "${paramsFile}"
    link=$(echo "${consentJson}" | jq -r '.value[0].link // empty' 2>/dev/null || echo "")
    if [[ -z "${link}" ]]; then
        echo -e "${RED}   list-consent-links returned no link; skipping${NC}"
        return
    fi

    echo -e "${CYAN}   opening browser for OAuth consent...${NC}"
    echo -e "${CYAN}   (if no tab opens, paste this URL manually:${NC}"
    echo -e "${CYAN}      ${link})${NC}"
    open_url "${link}"

    local deadline=$(($(date +%s) + 300))
    local lastStatus=""
    local s=""
    while [[ $(date +%s) -lt $deadline ]]; do
        s=$(az connector-namespace connection show \
            -g "${resourceGroupName}" --namespace "${connectorNamespaceName}" \
            -n "${connectionName}" --query "properties.overallStatus" -o tsv 2>/dev/null || echo "")
        if [[ "$s" != "$lastStatus" ]]; then
            echo -e "${CYAN}   status: ${s:-?}${NC}"
            lastStatus="$s"
        fi
        if [[ "$(echo "$s" | tr '[:upper:]' '[:lower:]')" == "connected" ]]; then
            echo -e "${GREEN}   ✓ ${connectionName} authenticated${NC}"
            return
        fi
        sleep 3
    done
    echo -e "${YELLOW}   timed out waiting for consent (5 min). Re-run azd up or this script when ready.${NC}"
}

authorize_connection "${connectorNamespaceConnectionName}"               "Office 365 Outlook (trigger + sender history + flag)"
authorize_connection "${connectorNamespaceTeamsConnectionName}"          "Teams (post triage card)"
authorize_connection "${connectorNamespaceOffice365usersConnectionName}" "Office 365 Users (IN-ORG badge + manager enrichment)"

echo ""
echo -e "${GREEN}✅ All connector connections authorized.${NC}"
echo ""
