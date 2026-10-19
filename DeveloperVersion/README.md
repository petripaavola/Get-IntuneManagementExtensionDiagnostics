# Current Development versions are published here

These version should fully work and they include new features.  
They may not be 100% polished so that's why it is still here in Developer Section

## Current Developer version is 3.4

📦 [**Download Developer version v3.4**](./Get-IntuneManagementExtensionDiagnostics.ps1)

Version 3.4 adds optional offline folder, ZIP and individual-log sources, streamed
CMTrace parsing, file coverage and rejected-record diagnostics. It also corrects
remediation lifecycle reporting and invalid verbose parameters (issues #13/#14).
Running without a source retains local IME log analysis and automatic ESP startup.

```powershell
.\Get-IntuneManagementExtensionDiagnostics.ps1 -InputPath "C:\Support\IME Logs.zip" -DoNotOpenReportAutomatically
```

The HTML report and its `.diagnostics.json` sidecar are written beside the script,
or to an existing directory supplied with `-ExportHTMLReportPath`. Both reports and
source logs may contain device or user information; keep them protected.

### Regression checks

On Windows, run the synthetic-fixture suite with a fresh writable output directory:

```powershell
.\tests\Test-IntuneManagementExtensionDiagnostics.ps1 -WorkPath "C:\Temp\IME-Regression"
```

The suite mocks Graph responses and local startup; it does not connect to a tenant,
read the workstation's actual IME logs or launch a browser. No device logs are
included in the repository.

The optional `-SampleZip` argument enables additional fixed-baseline checks against
the original development sample (11 files, 38,891 records, 34 timeline events).
It is not a generic assertion for arbitrary diagnostic archives, and that sample
is intentionally not published. Synthetic folder/ZIP equivalence is checked even
when the sample is omitted.
