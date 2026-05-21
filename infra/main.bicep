targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the the environment which is used to generate a short unique hash used in all resources.')
param environmentName string

@metadata({
  azd: {
    type: 'location'
  }
})
param location string

metadata name = 'Azure Functions Connectors Demo'
metadata description = 'Creates Azure Functions Flex Consumption and Connector Namespace resources'

@description('Id of the user identity to be used for testing and debugging. This is not required in production. Leave empty if not needed.')
@metadata({
  azd: {
    type: 'principalId'
  }
})
param userPrincipalId string  = deployer().objectId

@description('Name of the Azure Function that handles the Office 365 connector trigger.')
param office365FunctionName string = 'OnNewImportantEmailReceived'

@description('The Teams Team ID (groupId) to post notifications to.')
param teamsTeamId string

@description('The Teams Channel ID to post notifications to.')
param teamsChannelId string

@description('Optional. Comma-separated list of email addresses (mail or UPN) whose messages always count as important (e.g. your manager, skip-level, key stakeholders).')
param importantSenders string = ''

@description('Optional. Comma-separated list of email domains considered internal/in-org (e.g. "microsoft.com,contoso.com"). Senders whose domain matches will be looked up via the Office 365 Users connector for IN-ORG badging and profile enrichment. When empty, every sender is looked up.')
param internalDomains string = ''

var abbrs = loadJsonContent('./abbreviations.json')
var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = { 'azd-env-name': environmentName }

var functionAppName = '${abbrs.webSitesFunctions}${resourceToken}'
var functionAppPlanName = '${abbrs.webServerFarms}${resourceToken}'
var functionAppIdentityName = '${abbrs.managedIdentityUserAssignedIdentities}${resourceToken}'
var resourceGroupName = '${abbrs.resourcesResourceGroups}${environmentName}'
var storageAccountName = '${abbrs.storageStorageAccounts}${resourceToken}'
var logAnalyticsName = '${abbrs.operationalInsightsWorkspaces}${resourceToken}'
var appInsightsName = '${abbrs.insightsComponents}${resourceToken}'
var connectorNamespaceName = '${abbrs.connectorNamespaces}${resourceToken}'
var connectorNamespaceConnectionName = '${abbrs.connectorNamespacesConnections}${resourceToken}'
var connectorNamespaceTeamsConnectionName = '${abbrs.connectorNamespacesConnections}teams-${resourceToken}'
var connectorNamespaceOffice365usersConnectionName = '${abbrs.connectorNamespacesConnections}o365users-${resourceToken}'

var deploymentStorageContainerName = 'app-package-${take(functionAppName, 32)}-${take(toLower(uniqueString(functionAppName, environmentName)), 7)}'
var storageBlobDataOwner = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
var storageQueueDataContributor = '974c5e8b-45b9-4653-ba55-5f855dd0fb88'
var storageTableDataContributor = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
var MonitoringMetricsPublisherRoleId = '3913510d-42f4-4e42-8a64-420c390055eb'

resource resourceGroup 'Microsoft.Resources/resourceGroups@2025-04-01' = {
  name: resourceGroupName
  location: location
  tags: tags
}

module logAnalytics 'br/public:avm/res/operational-insights/workspace:0.15.0' = {
  name: '${uniqueString(deployment().name, location)}-loganalytics'
  scope: resourceGroup
  params: {
    name: logAnalyticsName
    location: location
    tags: tags
    dataRetention: 30
  }
}

module monitoring 'br/public:avm/res/insights/component:0.7.1' = {
  name: '${uniqueString(deployment().name, location)}-appinsights'
  scope: resourceGroup
  params: {
    name: appInsightsName
    location: location
    tags: tags
    workspaceResourceId: logAnalytics.outputs.resourceId
    disableLocalAuth: true
    roleAssignments: [
      {
        roleDefinitionIdOrName: MonitoringMetricsPublisherRoleId
        principalId: funcUserAssignedIdentity.outputs.principalId
        principalType: 'ServicePrincipal'
      }
      {
        roleDefinitionIdOrName: MonitoringMetricsPublisherRoleId
        principalId: userPrincipalId
        principalType: 'User'
      }
    ]
  }
}

module storageAccount 'br/public:avm/res/storage/storage-account:0.32.0' = {
  scope: resourceGroup
  name: storageAccountName
  params: {
    name: storageAccountName
    location: location
    tags: tags
    skuName: 'Standard_LRS'
    kind: 'StorageV2'
    allowBlobPublicAccess: false
    allowSharedKeyAccess: false
    publicNetworkAccess: 'Enabled'
    networkAcls: {
      defaultAction: 'Allow'
    } 
    minimumTlsVersion: 'TLS1_2'
    blobServices: {
      containers: [{name: deploymentStorageContainerName}]
    }
    roleAssignments:  [
      {
        roleDefinitionIdOrName: storageBlobDataOwner
        principalId: funcUserAssignedIdentity.outputs.principalId
        principalType: 'ServicePrincipal'
      }
      {
        roleDefinitionIdOrName: storageQueueDataContributor
        principalId: funcUserAssignedIdentity.outputs.principalId
        principalType: 'ServicePrincipal'
      }
      {
        roleDefinitionIdOrName: storageTableDataContributor
        principalId: funcUserAssignedIdentity.outputs.principalId
        principalType: 'ServicePrincipal'
      }
      {
        roleDefinitionIdOrName: storageBlobDataOwner
        principalId: userPrincipalId
        principalType: 'User'
      }
      {
        roleDefinitionIdOrName: storageQueueDataContributor
        principalId: userPrincipalId
        principalType: 'User'
      }
      {
        roleDefinitionIdOrName: storageTableDataContributor
        principalId: userPrincipalId
        principalType: 'User'
      }
    ]
  }
}



module funcUserAssignedIdentity 'br/public:avm/res/managed-identity/user-assigned-identity:0.5.0' = {
  name: 'funcUserAssignedIdentity'
  scope: resourceGroup
  params: {
    location: location
    tags: tags
    name: functionAppIdentityName
  }
}

// Function App Plan (Flex Consumption)
module functionAppPlan 'br/public:avm/res/web/serverfarm:0.7.0' = {
  scope: resourceGroup
  name: functionAppPlanName
  params: {
    name: functionAppPlanName
    location: location
    tags: tags
    skuName: 'FC1'
    reserved: true
  }
}

// Connector Namespace
module connectorNamespace './connectorNamespace.bicep' = {
  scope: resourceGroup
  name: connectorNamespaceName
  params: {
    name: connectorNamespaceName
    location: 'brazilsouth' // Connector Namespace features we need are only available in Brazil South as of now
    tags: tags
    connectionName: connectorNamespaceConnectionName
    connectorName: 'office365'
    teamsConnectionName: connectorNamespaceTeamsConnectionName
    office365usersConnectionName: connectorNamespaceOffice365usersConnectionName
    functionAppPrincipalId: funcUserAssignedIdentity.outputs.principalId
    userPrincipalId: userPrincipalId
  }
}

// Function App
module functionApp 'br/public:avm/res/web/site:0.22.0' = {
  scope: resourceGroup
  name: functionAppName
  params: {
    name: functionAppName
    location: location
    tags: union(tags, { 'azd-service-name': 'function-app' })
    kind: 'functionapp,linux'
    serverFarmResourceId: functionAppPlan.outputs.resourceId
    httpsOnly: true
    managedIdentities: {
      userAssignedResourceIds: [
        '${funcUserAssignedIdentity.outputs.resourceId}'
      ]
    }
    functionAppConfig: {
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${storageAccount.outputs.primaryBlobEndpoint}${deploymentStorageContainerName}'
          authentication: {
            type: 'UserAssignedIdentity'
            userAssignedIdentityResourceId: funcUserAssignedIdentity.outputs.resourceId 
          }
        }
      }
      scaleAndConcurrency: {
        instanceMemoryMB: 2048
        maximumInstanceCount: 100
      }
      runtime: {
        name: 'dotnet-isolated'
        version: '10.0'
      }
    }
    siteConfig: {
      alwaysOn: false
    }
    configs: [{
      name: 'appsettings'
        properties: {
          AzureWebJobsStorage__credential: 'managedidentity'
          AzureWebJobsStorage__clientId: funcUserAssignedIdentity.outputs.clientId
          AzureWebJobsStorage__accountName: storageAccount.outputs.name
          APPLICATIONINSIGHTS_AUTHENTICATION_STRING: 'ClientId=${funcUserAssignedIdentity.outputs.clientId};Authorization=AAD'
          APPLICATIONINSIGHTS_CONNECTION_STRING: monitoring.outputs.connectionString
          AZURE_CLIENT_ID: funcUserAssignedIdentity.outputs.clientId //Used by Open Telemetry managed identity
          TEAMS_CONNECTION_RUNTIME_URL: connectorNamespace.outputs.teamsConnectionRuntimeUrl
          OFFICE365_CONNECTION_RUNTIME_URL: connectorNamespace.outputs.office365ConnectionRuntimeUrl
          OFFICE365USERS_CONNECTION_RUNTIME_URL: connectorNamespace.outputs.office365usersConnectionRuntimeUrl
          TEAMS_TEAM_ID: teamsTeamId
          TEAMS_CHANNEL_ID: teamsChannelId
          IMPORTANT_SENDERS: importantSenders
          INTERNAL_DOMAINS: internalDomains
        }
      }]
  }
}


@description('The resource ID of the created Resource Group.')
output resourceGroupResourceId string = resourceGroup.id

@description('The name of the created Resource Group.')
output resourceGroupName string = resourceGroup.name

@description('The name of the created Function App.')
output functionAppName string = functionApp.outputs.name

@description('The default hostname of the created Function App.')
output functionAppDefaultHostname string = functionApp.outputs.defaultHostname

@description('The name of the created Connector Namespace.')
output connectorNamespaceName string = connectorNamespace.outputs.name

@description('The name of the created Connector Namespace Connection.')
output connectorNamespaceConnectionName string = connectorNamespace.outputs.connectionName

@description('The name of the created Teams Connector Namespace Connection.')
output connectorNamespaceTeamsConnectionName string = connectorNamespace.outputs.teamsConnectionName

@description('The name of the created Office 365 Users Connector Namespace Connection.')
output connectorNamespaceOffice365usersConnectionName string = connectorNamespace.outputs.office365usersConnectionName

@description('The name of the Function that handles the Office 365 connector trigger.')
output office365FunctionName string = office365FunctionName
