# Changelog

All notable changes to this project will be documented in this file.

The format follows [Conventional Commits](https://www.conventionalcommits.org/) and this project adheres to [Semantic Versioning](https://semver.org/). Release notes for each version are also generated from git history by the automation pipeline using the same conventional types (feat, fix, docs, refactor, test, etc.).

## [1.0.0] - (initial)

### Features

- Initial release of the RKSolutions PowerShell module.
- Cmdlets: Connect-RKGraph, Disconnect-RKGraph, Get-IntuneEnrollmentFlowsReport, Get-IntuneAnomaliesReport, Get-EntraAdminRolesReport, Get-M365LicenseAssignmentReport, Get-DeviceEvaluationContext, Get-CloudPCProvisioningPolicyGroupInfo.
- Connects to Microsoft Graph and generates HTML/CSV reports for Intune enrollment flows, anomalies, Entra admin roles, and M365 license assignment.
