# Cmdlet Reference

Summary of exported cmdlets and main parameters. Run `Get-Help <CmdletName> -Full` for full help.

## Connect-RKGraph

Establishes a Microsoft Graph session for use with RKSolutions report cmdlets.


| Parameter                 | Description                                                                                          |
| ------------------------- | ---------------------------------------------------------------------------------------------------- |
| **RequiredScopes**        | API permission scopes. Default includes scopes needed for all report cmdlets.                        |
| **TenantId**              | Tenant ID (optional for interactive; required for ClientSecret, Certificate, Identity, AccessToken). |
| **ClientId**              | App (client) ID for client secret or certificate auth.                                               |
| **ClientSecret**          | Client secret as **SecureString** (e.g. `ConvertTo-SecureString -String '...' -AsPlainText -Force`). |
| **CertificateThumbprint** | Certificate thumbprint for certificate auth.                                                         |
| **Identity**              | Use managed identity.                                                                                |
| **AccessToken**           | Access token as **SecureString**.                                                                     |
| **DebugMode**             | Enable debug output.                                                                                 |


**Example:** `Connect-RKGraph` (interactive); or `Connect-RKGraph -TenantId '...' -ClientId '...' -ClientSecret (ConvertTo-SecureString '...' -AsPlainText -Force)`.

---

## Disconnect-RKGraph

Disconnects from Microsoft Graph and clears the session.

**Example:** `Disconnect-RKGraph`

---

## Get-IntuneEnrollmentFlowsReport

Generates Intune assignment overview and/or device visualization report (HTML/CSV).


| Parameter                                            | Description                                                                         |
| ---------------------------------------------------- | ----------------------------------------------------------------------------------- |
| **AssignmentOverviewOnly**                           | Run assignment collection only (no device).                                         |
| **Device**                                           | Device display name, Intune device ID, or Entra device object ID for visualization. |
| **OutputPath**                                       | Output file path.                                                                   |
| **ExportToCsv**                                      | Also export CSV.                                                                    |
| **ExportFolder**                                     | Folder for exports.                                                                 |
| **DebugMode**                                        | Enable debug output.                                                                |

Connect first with **Connect-RKGraph**; this cmdlet uses the existing connection (no auth parameters).


**Example:** `Get-IntuneEnrollmentFlowsReport -AssignmentOverviewOnly`

---

## Get-IntuneAnomaliesReport

Generates Intune anomalies report.


| Parameter    | Description                    |
| ------------ | ------------------------------ |
| **SendEmail** | Send report by email.          |
| **Recipient** | Email recipient(s).            |
| **From**      | From address.                  |
| **ExportPath** | Output file path.            |
| **DebugMode** | Enable debug output.          |

Connect first with **Connect-RKGraph**; this cmdlet uses the existing connection (no auth parameters).


**Example:** `Get-IntuneAnomaliesReport`

---

## Get-EntraAdminRolesReport

Generates Entra admin roles report.


| Parameter    | Description                    |
| ------------ | ------------------------------ |
| **SendEmail** | Send report by email.          |
| **Recipient** | Email recipient(s).            |
| **From**      | From address.                  |
| **ExportPath** | Output file path.            |
| **DebugMode** | Enable debug output.          |

Connect first with **Connect-RKGraph**; this cmdlet uses the existing connection (no auth parameters).


**Example:** `Get-EntraAdminRolesReport`

---

## Get-M365LicenseAssignmentReport

Generates M365 license assignment report.


| Parameter    | Description                    |
| ------------ | ------------------------------ |
| **SendEmail** | Send report by email.          |
| **Recipient** | Email recipient(s).            |
| **From**      | From address.                  |
| **ExportPath** | Output file path.            |
| **DebugMode** | Enable debug output.          |

Connect first with **Connect-RKGraph**; this cmdlet uses the existing connection (no auth parameters).


**Example:** `Get-M365LicenseAssignmentReport`

---

## Get-CustomSecurityAttributesReport

Generates an interactive HTML report of custom security attributes across users, devices, and enterprise applications. Auto-discovers attribute sets or accepts a specific set.


| Parameter       | Description                                                                  |
| --------------- | ---------------------------------------------------------------------------- |
| **AttributeSet** | Specific attribute set to report on. If omitted, discovers and reports all. |
| **SendEmail**   | Send report by email.                                                        |
| **Recipient**   | Email recipient(s).                                                          |
| **From**        | From address.                                                                |
| **ExportPath**  | Output file path.                                                            |
| **DebugMode**   | Enable debug output.                                                         |

Connect first with **Connect-RKGraph**; this cmdlet uses the existing connection (no auth parameters).


**Examples:**
- `Get-CustomSecurityAttributesReport` — auto-discover all attribute sets
- `Get-CustomSecurityAttributesReport -AttributeSet "ComplianceData"` — specific set only
- `Get-CustomSecurityAttributesReport -SendEmail -Recipient "admin@contoso.com"` — generate and email

---

For full parameter sets and examples, run `Get-Help <CmdletName> -Full` in PowerShell.