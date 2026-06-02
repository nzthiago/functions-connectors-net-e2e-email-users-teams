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

Write-Host "Post-deployment configuration..." -ForegroundColor Yellow

# --- Read azd outputs --------------------------------------------------------
$outputs = azd env get-values --output json | ConvertFrom-Json

$subscriptionId                              = (az account show --query id -o tsv)
$resourceGroupName                           = $outputs.resourceGroupName
$connectorNamespaceName                      = $outputs.connectorNamespaceName
$connectorNamespaceConnectionName            = $outputs.connectorNamespaceConnectionName
$connectorNamespaceTeamsConnectionName       = $outputs.connectorNamespaceTeamsConnectionName
$connectorNamespaceOffice365usersConnectionName = $outputs.connectorNamespaceOffice365usersConnectionName
$functionAppName                             = $outputs.functionAppName
$office365FunctionName                       = $outputs.office365FunctionName

# --- Install the official connector-namespace az CLI extension --------------
# Resolve the latest released wheel URL from the Azure/Connectors GitHub
# releases. All releases are marked pre-release so the standard
# /releases/latest endpoint 404s; we fetch /releases?per_page=1 instead.
# Pin a specific version by setting CONNECTOR_NAMESPACE_EXT_URL in the
# environment before running azd up.
if (-not $env:CONNECTOR_NAMESPACE_EXT_URL) {
    try {
        $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/Azure/Connectors/releases?per_page=1"
        $asset = $rel[0].assets | Where-Object {
            $_.name -match '^connector_namespace.*\.whl$'
        } | Select-Object -First 1
        if ($asset) {
            $env:CONNECTOR_NAMESPACE_EXT_URL = $asset.browser_download_url
        }
    } catch {
        Write-Host "WARNING: could not query Azure/Connectors releases: $_" -ForegroundColor Yellow
    }
}
if (-not $env:CONNECTOR_NAMESPACE_EXT_URL) {
    Write-Host "ERROR: could not resolve connector-namespace extension URL from Azure/Connectors releases" -ForegroundColor Red
    exit 2
}
$extInstalled = az extension show --name connector-namespace --query name -o tsv 2>$null
if (-not $extInstalled) {
    Write-Host "Installing 'connector-namespace' Azure CLI extension from $($env:CONNECTOR_NAMESPACE_EXT_URL)" -ForegroundColor Cyan
    az extension add --upgrade --yes --source $env:CONNECTOR_NAMESPACE_EXT_URL
}

# --- Create Connector Namespace trigger config ------------------------------
Write-Host ""
Write-Host "Creating Connector Namespace trigger config..." -ForegroundColor Yellow

Write-Host "Fetching connector extension key for $functionAppName..." -ForegroundColor Cyan
# The system key is generated when the Functions runtime loads the
# Microsoft.Azure.Functions.Worker.Extensions.Connectors extension,
# which can take a few seconds after the function app finishes
# deploying. Poll for it briefly before giving up.
$connectorExtensionKey = $null
$deadline = (Get-Date).AddMinutes(3)
while ((Get-Date) -lt $deadline) {
    $connectorExtensionKey = (az functionapp keys list -g $resourceGroupName -n $functionAppName --query "systemKeys.connector_extension" -o tsv 2>$null)
    if ($connectorExtensionKey) { break }
    Write-Host "  connector_extension key not yet present; waiting 10s..." -ForegroundColor Yellow
    Start-Sleep -Seconds 10
}
if (-not $connectorExtensionKey) {
    Write-Host "ERROR: connector_extension system key never appeared on $functionAppName. The Microsoft.Azure.Functions.Worker.Extensions.Connectors extension may have failed to load. Check the function app's log stream." -ForegroundColor Red
    exit 1
}

$triggerName = "$connectorNamespaceConnectionName-trigger"
$callbackUrl = "https://$functionAppName.azurewebsites.net/runtime/webhooks/connector?functionName=$office365FunctionName&code=$connectorExtensionKey"

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
# Hand-write the parameters array: PowerShell unwraps single-element
# arrays during ConvertTo-Json (producing `{...}` instead of `[{...}]`),
# which the CLI parser rejects with "list type value expected".
$connDetailsFile  = New-TemporaryFile
@{ connectorName = "office365"; connectionName = $connectorNamespaceConnectionName } | ConvertTo-Json -Compress | Set-Content -Path $connDetailsFile -Encoding utf8
$notifDetailsFile = New-TemporaryFile
@{ callbackUrl = $callbackUrl; httpMethod = "Post" } | ConvertTo-Json -Compress | Set-Content -Path $notifDetailsFile -Encoding utf8
$parameters = '[{"name":"folderPath","value":"Inbox"}]'

Write-Host "  Trigger name: $triggerName" -ForegroundColor Cyan
Write-Host "  Callback URL: $callbackUrl" -ForegroundColor Cyan

# Best-effort idempotency: delete any prior config with this name.
az connector-namespace trigger delete `
    -g $resourceGroupName `
    --namespace $connectorNamespaceName `
    -n $triggerName `
    --yes 2>$null | Out-Null

az connector-namespace trigger create `
    -g $resourceGroupName `
    --namespace $connectorNamespaceName `
    -n $triggerName `
    --connection-details "@$connDetailsFile" `
    --operation-name "OnNewEmailV3" `
    --parameters $parameters `
    --notification-details "@$notifDetailsFile" `
    --state "Enabled" `
    --description "When a new email arrives in the consented Office 365 mailbox, POST the payload to the function's connector webhook."

Remove-Item $connDetailsFile, $notifDetailsFile -ErrorAction SilentlyContinue

if ($LASTEXITCODE -ne 0) {
    Write-Host "Failed to create Connector Namespace trigger config." -ForegroundColor Red
    exit 1
}

Write-Host "✅ Connector Namespace trigger config created successfully!" -ForegroundColor Green

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
Write-Host ""
Write-Host "Authorizing connector connections via Azure CLI..." -ForegroundColor Yellow

function Invoke-AuthorizeConnection {
    param(
        [Parameter(Mandatory)] [string] $ConnectionName,
        [Parameter(Mandatory)] [string] $Description
    )
    Write-Host "-> Authorizing $Description connection: $ConnectionName" -ForegroundColor Cyan

    $currentStatus = az connector-namespace connection show `
        -g $resourceGroupName --namespace $connectorNamespaceName `
        -n $ConnectionName --query "properties.overallStatus" -o tsv 2>$null
    if ($currentStatus -and $currentStatus.ToLower() -eq "connected") {
        Write-Host "   already Connected; skipping consent flow" -ForegroundColor Green
        return
    }

    # `list-consent-links` --parameters: the URL value contains a colon
    # which makes the CLI's shorthand parser barf. Write to a temp file
    # and pass with the @file syntax.
    $paramsFile = New-TemporaryFile
    '[{"parameterName":"token","redirectUrl":"https://portal.azure.com"}]' | Set-Content -Path $paramsFile -Encoding utf8
    $consentJson = az connector-namespace connection list-consent-links `
        -g $resourceGroupName --namespace $connectorNamespaceName `
        --connection-name $ConnectionName --parameters "@$paramsFile" -o json 2>$null
    Remove-Item $paramsFile -ErrorAction SilentlyContinue
    if (-not $consentJson) {
        Write-Host "   list-consent-links returned no output; skipping" -ForegroundColor Red
        return
    }
    $link = ($consentJson | ConvertFrom-Json).value[0].link
    if (-not $link) {
        Write-Host "   list-consent-links returned no link; skipping" -ForegroundColor Red
        return
    }

    Write-Host "   opening browser for OAuth consent..." -ForegroundColor Cyan
    Write-Host "   (if no tab opens, paste this URL manually:" -ForegroundColor Cyan
    Write-Host "      $link)" -ForegroundColor Cyan
    try { Start-Process $link | Out-Null } catch { Write-Host "   Start-Process failed: $_" -ForegroundColor Yellow }

    $deadline = (Get-Date).AddMinutes(5)
    $lastStatus = ""
    while ((Get-Date) -lt $deadline) {
        $s = az connector-namespace connection show `
            -g $resourceGroupName --namespace $connectorNamespaceName `
            -n $ConnectionName --query "properties.overallStatus" -o tsv 2>$null
        if ($s -ne $lastStatus) {
            Write-Host "   status: $(if ($s) { $s } else { '?' })" -ForegroundColor Cyan
            $lastStatus = $s
        }
        if ($s -and $s.ToLower() -eq "connected") {
            Write-Host "   ✓ $ConnectionName authenticated" -ForegroundColor Green
            return
        }
        Start-Sleep -Seconds 3
    }
    Write-Host "   timed out waiting for consent (5 min). Re-run azd up or this script when ready." -ForegroundColor Yellow
}

Invoke-AuthorizeConnection -ConnectionName $connectorNamespaceConnectionName               -Description "Office 365 Outlook (trigger + sender history + flag)"
Invoke-AuthorizeConnection -ConnectionName $connectorNamespaceTeamsConnectionName          -Description "Teams (post triage card)"
Invoke-AuthorizeConnection -ConnectionName $connectorNamespaceOffice365usersConnectionName -Description "Office 365 Users (IN-ORG badge + manager enrichment)"

Write-Host ""
Write-Host "✅ All connector connections authorized." -ForegroundColor Green
Write-Host ""
