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

apiUrl="https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${resourceGroupName}/providers/Microsoft.Web/connectorNamespaces/${connectorNamespaceName}/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

body=$(cat <<EOF
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
EOF
)

echo -e "${CYAN}  API URL: ${apiUrl}${NC}"
echo -e "${CYAN}  Callback URL: ${callbackUrl}${NC}"

az rest --method PUT --url "${apiUrl}" --body "${body}"

echo -e "${GREEN}✅ Connector Namespace trigger config created successfully!${NC}"

echo ""
echo -e "${YELLOW}╔══════════════════════════════════════════════════════════════════════╗${NC}"
echo -e "${YELLOW}║  ⚠️  IMPORTANT: Authorize the Connector Connections                 ║${NC}"
echo -e "${YELLOW}╠══════════════════════════════════════════════════════════════════════╣${NC}"
echo -e "${YELLOW}║                                                                      ║${NC}"
echo -e "${YELLOW}║  Before testing, you must authorize all connector connections:       ║${NC}"
echo -e "${YELLOW}║                                                                      ║${NC}"
echo -e "${YELLOW}║  1. Open the Azure Portal: https://portal.azure.com                  ║${NC}"
echo -e "${YELLOW}║  2. Navigate to Resource Group: ${resourceGroupName}${NC}"
echo -e "${YELLOW}║  3. Open the Connector Namespace resource: ${connectorNamespaceName}${NC}"
echo -e "${YELLOW}║  4. Go to Connections → authorize the Office 365 connection:         ║${NC}"
echo -e "${YELLOW}║     ${connectorNamespaceConnectionName}                                ║${NC}"
echo -e "${YELLOW}║     (used by both the trigger and for sender history + flag actions) ║${NC}"
echo -e "${YELLOW}║  5. Authorize the Teams connection:                                  ║${NC}"
echo -e "${YELLOW}║     ${connectorNamespaceTeamsConnectionName}                           ║${NC}"
echo -e "${YELLOW}║  6. Authorize the Office 365 Users connection:                       ║${NC}"
echo -e "${YELLOW}║     ${connectorNamespaceOffice365usersConnectionName}                  ║${NC}"
echo -e "${YELLOW}║     Used to look up the sender's M365 profile (UserProfileAsync +    ║${NC}"
echo -e "${YELLOW}║     ManagerAsync) for IN-ORG badging and card enrichment.            ║${NC}"
echo -e "${YELLOW}║                                                                      ║${NC}"
echo -e "${YELLOW}║  The trigger will NOT fire until Office 365 connection is authorized. ║${NC}"
echo -e "${YELLOW}║  Teams notifications require the Teams connection to be authorized.   ║${NC}"
echo -e "${YELLOW}║  IN-ORG badging requires the Office 365 Users connection.             ║${NC}"
echo -e "${YELLOW}╚══════════════════════════════════════════════════════════════════════╝${NC}"
echo ""
