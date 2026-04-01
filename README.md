# HelloID-Conn-SA-Full-Teams-AddChannel

| :information_source: Information                                                                                                                                                                                                                                                                                                                                                          |
|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible for acquiring the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements. |

## Description
_HelloID-Conn-SA-Full-Teams-AddChannel_ is a template designed for use with HelloID Service Automation (SA) Delegated Forms. It can be imported into HelloID and customized according to your requirements.

By using this delegated form , you can create Microsoft Teams channels through Microsoft Graph. The delegated form supports the following flow:
1. Search for and select an existing Microsoft Team
2. Enter channel details (name and description)
3. Choose between standard or private channel
4. For private channels, select one or more owners
5. Create the channel in Microsoft 365

## Getting started
### Requirements

- **Microsoft Entra application registration (certificate-based)**:
  The connector authenticates to Microsoft Graph using a certificate (client credentials flow).
- **Microsoft Graph application permissions**:
  Configure and grant admin consent for the following minimal application permissions:
  - `Channel.Create.Group`
  - `GroupMember.Read.All`

### Connection settings

The following user-defined variables are used by the connector.

| Setting                        | Description                                                                | Mandatory |
|--------------------------------|----------------------------------------------------------------------------|-----------|
| EntraIdTenantId                | Microsoft Entra tenant ID                                                  | Yes       |
| EntraIdAppId                   | Application (client) ID of the app registration                            | Yes       |
| EntraIdCertificateBase64String | Base64 encoded certificate (including private key) used for authentication | Yes       |
| EntraIdCertificatePassword     | Password for the certificate                                               | Yes       |

## Remarks

### Microsoft Graph Query Behavior
- `ConsistencyLevel: eventual` is added on Graph requests where advanced query capabilities are used (for example filtering and searching result sets). This allows Microsoft Graph to evaluate those queries correctly and consistently, especially when data has just changed.

### Private vs. Standard Channels
- Standard channels are created without owners defined at creation time.
- Private channels can have owners, who are included in the channel creation payload for immediate membership.

## Development resources

### API endpoints

The following endpoints are used by the connector.

| Endpoint                                                    | Description                                                             |
|-------------------------------------------------------------|-------------------------------------------------------------------------|
| `https://login.microsoftonline.com/{tenantId}/oauth2/token` | Retrieve OAuth2 access token using certificate-based client credentials |
| `https://graph.microsoft.com/v1.0/groups`                   | Search Teams-enabled groups                                             |
| `https://graph.microsoft.com/v1.0/teams/{teamId}/channels`  | Create channel in the selected team                                     |

### API documentation

- https://learn.microsoft.com/graph/api/overview
- https://learn.microsoft.com/graph/api/channel-post
- https://learn.microsoft.com/graph/api/group-list

## Getting help
> :bulb: **Tip:**  
> _For more information on Delegated Forms, please refer to our [documentation](https://docs.helloid.com/en/service-automation/delegated-forms.html) pages_.

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/
