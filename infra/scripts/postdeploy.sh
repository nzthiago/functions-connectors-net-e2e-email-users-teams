#!/bin/bash
set -e

# Colors for output
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
WHITE='\033[1;37m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${YELLOW}Post-deployment configuration...${NC}"

# Get outputs from azd
outputs=$(azd env get-values --output json)

if command -v jq &> /dev/null; then
    subscriptionId=$(echo "$outputs" | jq -r '.AZURE_SUBSCRIPTION_ID')
    resourceGroupName=$(echo "$outputs" | jq -r '.resourceGroupName')
    connectorNamespaceName=$(echo "$outputs" | jq -r '.connectorNamespaceName')
    connectorNamespaceConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceConnectionName')
    connectorNamespaceTeamsConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceTeamsConnectionName')
    connectorNamespaceOffice365usersConnectionName=$(echo "$outputs" | jq -r '.connectorNamespaceOffice365usersConnectionName')
    functionAppName=$(echo "$outputs" | jq -r '.functionAppName')
    office365FunctionName=$(echo "$outputs" | jq -r '.office365FunctionName')
else
    echo -e "${RED}Error: jq is required for this script. Please install jq.${NC}"
    exit 1
fi

# --- Create Connector Namespace trigger config ---
echo -e "${YELLOW}Creating Connector Namespace trigger config...${NC}"

# Fetch the connector extension system key
echo -e "${CYAN}Fetching connector extension key for ${functionAppName}...${NC}"
connectorExtensionKey=$(az functionapp keys list -g "${resourceGroupName}" -n "${functionAppName}" --query "systemKeys.connector_extension" -o tsv)

triggerName="${connectorNamespaceConnectionName}-trigger"
callbackUrl="https://${functionAppName}.azurewebsites.net/runtime/webhooks/connector?functionName=${office365FunctionName}&code=${connectorExtensionKey}"

apiUrl="https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/connectorGateways/${connectorNamespaceName}/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

body=$(cat <<JSON
{
  "properties": {
    "description": "Office 365 Outlook trigger config",
    "connectionDetails": {
      "connectorName": "office365",
      "connectionName": "${connectorNamespaceConnectionName}"
    },
    "operationName": "OnNewEmailV3",
    "parameters": [
      {
        "name": "folderPath",
        "value": "Inbox"
      }
    ],
    "notificationDetails": {
      "callbackUrl": "${callbackUrl}"
    }
  }
}
JSON
)

echo -e "${CYAN}  API URL: ${apiUrl}${NC}"
echo -e "${CYAN}  Callback URL: ${callbackUrl}${NC}"

az rest --method PUT --url "${apiUrl}" --body "${body}"

echo -e "${GREEN}✅ Connector Namespace trigger config created successfully!${NC}"

# --- Authorize the connector connections via Azure CLI ---
# Portal authorization UX is not yet available, so we drive the OAuth consent flow
# through the `connector-namespace` CLI extension. Each call opens a browser tab
# for the signed-in user to consent to the connection.
echo ""
echo -e "${YELLOW}Authorizing connector connections via Azure CLI...${NC}"

if ! az extension show --name connector-namespace >/dev/null 2>&1; then
  echo -e "${CYAN}Installing 'connector-namespace' Azure CLI extension...${NC}"
  az extension add \
    --source https://github.com/anthonychu/azure-cli-extensions/releases/download/connector-namespace-0.1.0/connector_namespace-0.1.0-py2.py3-none-any.whl \
    --yes
fi

authorize_connection() {
  local connectionName="$1"
  local description="$2"
  echo -e "${CYAN}-> Authorizing ${description} connection: ${connectionName}${NC}"
  echo -e "${CYAN}   A browser tab will open for OAuth consent. Sign in with the account that should back this connection.${NC}"
  az connector-namespace connection authorize \
    --resource-group "${resourceGroupName}" \
    --namespace-name "${connectorNamespaceName}" \
    --name "${connectionName}"
}

authorize_connection "${connectorNamespaceConnectionName}"               "Office 365 Outlook (trigger + sender history + flag)"
authorize_connection "${connectorNamespaceTeamsConnectionName}"          "Teams (post triage card)"
authorize_connection "${connectorNamespaceOffice365usersConnectionName}" "Office 365 Users (IN-ORG badge + manager enrichment)"

echo ""
echo -e "${GREEN}✅ All connector connections authorized.${NC}"
echo ""
