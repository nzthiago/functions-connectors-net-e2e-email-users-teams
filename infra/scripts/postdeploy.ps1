Write-Host "Post-deployment configuration..." -ForegroundColor Yellow

# Get outputs from azd
$outputs = azd env get-values --output json | ConvertFrom-Json

$subscriptionId = (az account show --query id -o tsv)
$resourceGroupName = $outputs.resourceGroupName
$connectorNamespaceName = $outputs.connectorNamespaceName
$connectorNamespaceConnectionName = $outputs.connectorNamespaceConnectionName
$connectorNamespaceTeamsConnectionName = $outputs.connectorNamespaceTeamsConnectionName
$connectorNamespaceOffice365usersConnectionName = $outputs.connectorNamespaceOffice365usersConnectionName
$functionAppName = $outputs.functionAppName
$office365FunctionName = $outputs.office365FunctionName

# --- Create Connector Namespace trigger config ---
Write-Host "Creating Connector Namespace trigger config..." -ForegroundColor Yellow

# Fetch the connector extension system key
Write-Host "Fetching connector extension key for $functionAppName..." -ForegroundColor Cyan
$connectorExtensionKey = (az functionapp keys list -g $resourceGroupName -n $functionAppName --query "systemKeys.connector_extension" -o tsv)

$triggerName = "$connectorNamespaceConnectionName-trigger"
$callbackUrl = "https://$functionAppName.azurewebsites.net/runtime/webhooks/connector?functionName=$office365FunctionName&code=$connectorExtensionKey"

$apiUrl = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Web/connectorNamespaces/$connectorNamespaceName/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

$body = @{
  properties = @{
    description = "Office 365 Outlook trigger config"
    connectionDetails = @{
      connectorName = "office365"
      connectionName = $connectorNamespaceConnectionName
    }
    operationName = "OnNewEmailV3"
    parameters = @(
      @{ name = "folderPath"; value = "Inbox" }
      @{ name = "fetchOnlyWithAttachment"; value = "false" }
      @{ name = "includeAttachments"; value = "false" }
    )
    notificationDetails = @{
      callbackUrl = $callbackUrl
    }
  }
} | ConvertTo-Json -Depth 5

$bodyFile = [System.IO.Path]::GetTempFileName()
$body | Out-File -FilePath $bodyFile -Encoding utf8

Write-Host "  API URL: $apiUrl" -ForegroundColor Cyan
Write-Host "  Callback URL: $callbackUrl" -ForegroundColor Cyan

az rest --method PUT --url $apiUrl --body "@$bodyFile" --headers "Content-Type=application/json"
Remove-Item $bodyFile -ErrorAction SilentlyContinue

if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to create Connector Namespace trigger config." -ForegroundColor Red
    exit 1
}

Write-Host "✅ Connector Namespace trigger config created successfully!" -ForegroundColor Green

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
Write-Host "║  ⚠️  IMPORTANT: Authorize the Connector Connections                  ║" -ForegroundColor Yellow
Write-Host "╠══════════════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  Before testing, you must authorize all connector connections:       ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  1. Open the Azure Portal: https://portal.azure.com                  ║" -ForegroundColor Yellow
Write-Host "║  2. Navigate to Resource Group: $resourceGroupName" -ForegroundColor Yellow
Write-Host "║  3. Open the Connector Namespace resource: $connectorNamespaceName" -ForegroundColor Yellow
Write-Host "║  4. Go to Connections -> authorize the Office 365 connection:         ║" -ForegroundColor Yellow
Write-Host "║     $connectorNamespaceConnectionName" -ForegroundColor Yellow
Write-Host "║     (used by both the trigger and for sender history + flag actions) ║" -ForegroundColor Yellow
Write-Host "║  5. Authorize the Teams connection:                                  ║" -ForegroundColor Yellow
Write-Host "║     $connectorNamespaceTeamsConnectionName" -ForegroundColor Yellow
Write-Host "║  6. Authorize the Office 365 Users connection:                       ║" -ForegroundColor Yellow
Write-Host "║     $connectorNamespaceOffice365usersConnectionName" -ForegroundColor Yellow
Write-Host "║     Used to look up the sender's M365 profile (UserProfileAsync +    ║" -ForegroundColor Yellow
Write-Host "║     ManagerAsync) for IN-ORG badging and card enrichment.            ║" -ForegroundColor Yellow
Write-Host "║                                                                      ║" -ForegroundColor Yellow
Write-Host "║  The trigger will NOT fire until Office 365 connection is authorized. ║" -ForegroundColor Yellow
Write-Host "║  Teams notifications require the Teams connection to be authorized.   ║" -ForegroundColor Yellow
Write-Host "║  IN-ORG badging requires the Office 365 Users connection.             ║" -ForegroundColor Yellow
Write-Host "╚══════════════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
Write-Host ""
