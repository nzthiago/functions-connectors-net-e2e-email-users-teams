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

$apiUrl = "https://management.azure.com/subscriptions/$subscriptionId/resourceGroups/$resourceGroupName/providers/Microsoft.Web/connectorGateways/$connectorNamespaceName/triggerconfigs/${triggerName}?api-version=2026-05-01-preview"

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

# --- Authorize the connector connections via Azure CLI ---
# Portal authorization UX is not yet available, so we drive the OAuth consent flow
# through the `connector-namespace` CLI extension. Each call opens a browser tab
# for the signed-in user to consent to the connection.
Write-Host ""
Write-Host "Authorizing connector connections via Azure CLI..." -ForegroundColor Yellow

$extInstalled = az extension show --name connector-namespace 2>$null
if (-not $extInstalled) {
    Write-Host "Installing 'connector-namespace' Azure CLI extension..." -ForegroundColor Cyan
    az extension add `
        --source https://github.com/anthonychu/azure-cli-extensions/releases/download/connector-namespace-0.1.0/connector_namespace-0.1.0-py2.py3-none-any.whl `
        --yes
}

function Invoke-AuthorizeConnection {
    param(
        [Parameter(Mandatory)] [string] $ConnectionName,
        [Parameter(Mandatory)] [string] $Description
    )
    Write-Host "-> Authorizing $Description connection: $ConnectionName" -ForegroundColor Cyan
    Write-Host "   A browser tab will open for OAuth consent. Sign in with the account that should back this connection." -ForegroundColor Cyan
    az connector-namespace connection authorize `
        --resource-group $resourceGroupName `
        --namespace-name $connectorNamespaceName `
        --name $ConnectionName
}

Invoke-AuthorizeConnection -ConnectionName $connectorNamespaceConnectionName               -Description "Office 365 Outlook (trigger + sender history + flag)"
Invoke-AuthorizeConnection -ConnectionName $connectorNamespaceTeamsConnectionName          -Description "Teams (post triage card)"
Invoke-AuthorizeConnection -ConnectionName $connectorNamespaceOffice365usersConnectionName -Description "Office 365 Users (IN-ORG badge + manager enrichment)"

Write-Host ""
Write-Host "✅ All connector connections authorized." -ForegroundColor Green
Write-Host ""
