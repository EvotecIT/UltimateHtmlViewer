# Local-Only Workspace (Not Committed)

Use this folder for operator-specific files that should never be shared in git, for example:

- tenant/runbook notes
- local test URLs
- temporary deployment snippets
- personal helper scripts

Recommended pattern:

1. Keep reusable templates in `scripts/examples/`.
2. Copy template locally into `ignore/`.
3. Fill in your tenant/site values locally.

Example:

```powershell
Copy-Item .\scripts\examples\UHV.LocalProfile.example.ps1 .\ignore\UHV.LocalProfile.ps1
```

Then load it before running scripts:

```powershell
. .\ignore\UHV.LocalProfile.ps1
```
