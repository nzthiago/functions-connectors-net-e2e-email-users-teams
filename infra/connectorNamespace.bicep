param name string
param location string
param tags object = {}
param connectionName string = ''
param connectorName string = ''
param teamsConnectionName string = ''
param office365usersConnectionName string = ''
param functionAppPrincipalId string = ''
@description('Optional. AAD object id of a user (typically the deployer) to also grant access to the Teams, Office 365, and Office 365 Users connections, so the same code can be debugged locally with `az login` credentials.')
param userPrincipalId string = ''
param tenantId string = tenant().tenantId

resource connectorNamespace 'Microsoft.Web/connectorGateways@2026-05-01-preview' = {
  name: name
  location: location
  tags: tags
  identity: {
    type: 'SystemAssigned'
  }
}

resource connectorNamespaceConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = if (!empty(connectionName)) {
  parent: connectorNamespace
  name: connectionName
  properties: {
    connectorName: connectorName
  }
}

// Allow the function's managed identity to call the Office 365 connection at
// runtime — used both for the trigger callback and for the Office365Client SDK
// calls (sender history + flag).
resource office365ConnectionAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(connectionName) && !empty(functionAppPrincipalId)) {
  parent: connectorNamespaceConnection
  name: 'functionapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: functionAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource office365ConnectionUserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(connectionName) && !empty(userPrincipalId)) {
  parent: connectorNamespaceConnection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource teamsConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = if (!empty(teamsConnectionName)) {
  parent: connectorNamespace
  name: teamsConnectionName
  properties: {
    connectorName: 'teams'
  }
}

resource teamsConnectionAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(teamsConnectionName) && !empty(functionAppPrincipalId)) {
  parent: teamsConnection
  name: 'functionapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: functionAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource teamsConnectionUserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(teamsConnectionName) && !empty(userPrincipalId)) {
  parent: teamsConnection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource office365usersConnection 'Microsoft.Web/connectorGateways/connections@2026-05-01-preview' = if (!empty(office365usersConnectionName)) {
  parent: connectorNamespace
  name: office365usersConnectionName
  properties: {
    connectorName: 'office365users'
  }
}

resource office365usersConnectionAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(office365usersConnectionName) && !empty(functionAppPrincipalId)) {
  parent: office365usersConnection
  name: 'functionapp-msi'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: functionAppPrincipalId
        tenantId: tenantId
      }
    }
  }
}

resource office365usersConnectionUserAccessPolicy 'Microsoft.Web/connectorGateways/connections/accessPolicies@2026-05-01-preview' = if (!empty(office365usersConnectionName) && !empty(userPrincipalId)) {
  parent: office365usersConnection
  name: 'dev-user'
  properties: {
    principal: {
      type: 'ActiveDirectory'
      identity: {
        objectId: userPrincipalId
        tenantId: tenantId
      }
    }
  }
}

@description('The resource ID of the Connector Namespace.')
output resourceId string = connectorNamespace.id

@description('The name of the Connector Namespace.')
output name string = connectorNamespace.name

@description('The resource ID of the Connector Namespace Connection.')
output connectionResourceId string = !empty(connectionName) ? connectorNamespaceConnection.id : ''

@description('The name of the Connector Namespace Connection.')
output connectionName string = !empty(connectionName) ? connectorNamespaceConnection.name : ''

@description('The connection runtime URL for the Office 365 Outlook connection.')
output office365ConnectionRuntimeUrl string = !empty(connectionName) ? connectorNamespaceConnection.properties.connectionRuntimeUrl : ''

@description('The name of the Teams Connector Namespace Connection.')
output teamsConnectionName string = !empty(teamsConnectionName) ? teamsConnection.name : ''

@description('The connection runtime URL for the Teams connection.')
output teamsConnectionRuntimeUrl string = !empty(teamsConnectionName) ? teamsConnection.properties.connectionRuntimeUrl : ''

@description('The name of the Office 365 Users Connector Namespace Connection.')
output office365usersConnectionName string = !empty(office365usersConnectionName) ? office365usersConnection.name : ''

@description('The connection runtime URL for the Office 365 Users connection.')
output office365usersConnectionRuntimeUrl string = !empty(office365usersConnectionName) ? office365usersConnection.properties.connectionRuntimeUrl : ''
