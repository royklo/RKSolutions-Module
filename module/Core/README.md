# Core

The files in this folder are copied into the RKSolutions Maester pack (royklo/maester-pack-rksolutions) by
tools/Sync-Core.ps1. They follow the Maester pack rules:

- only function definitions at the top level; script-scope values are created by Initialize-RKIntuneCoreState;
- Microsoft Graph is only read, and only through Invoke-RKCoreGraphRequest and Invoke-RKCoreGraphBatch;
- no other host, no Invoke-WebRequest or Invoke-RestMethod, no Invoke-Expression, no scriptblock variables
  called with &, no Add-Type, no computed command names;
- no variables from the caller's scope.

Run the pack repo's tools/Test-Pack.ps1 -Scan after a change; it must print "SCAN OK".
