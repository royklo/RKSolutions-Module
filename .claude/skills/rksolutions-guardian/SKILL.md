---
name: rksolutions-guardian
description: >
  Enforces conventions for the RKSolutions PowerShell module (Microsoft Graph reporting for
  Intune, Entra, M365). Use when changing cmdlets, parameters, permissions, or Graph API calls.
  Triggers on any code change in the RKSolutions-Module project.
---

# RKSolutions Guardian

> **Stub skill.** This skill was created during the Cursor → Claude migration. Expand it as
> conventions solidify. Track findings in `.claude/FINDINGS.md`.

## Triggers

Activate when working on the RKSolutions-Module project and the task involves:
- Changing cmdlets or parameters
- Updating Graph API calls or permissions
- Reviewing `-RequiredScopes` parameters
- Auth changes (Connect-RKGraph)

## Key Conventions

- Per-cmdlet `-RequiredScopes` parameter for Graph permissions
- Permissions documented in `docs/PERMISSIONS.md`
- Auth via `Connect-RKGraph`

## Findings

Track issues and fixes in `.claude/FINDINGS.md`.

## Key Files

| Purpose | Path |
|---------|------|
| Findings | `.claude/FINDINGS.md` |
| Permissions docs | `docs/PERMISSIONS.md` |
