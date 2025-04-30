# TFSTestRunExport

Export Test Runs from Team Foundation Server (or Microsoft Test Manager) to Excel.
NOTE: This is a fork of an existing repo.

## Prerequisites
Install these using Package Manager Console:
```
Install-Package Microsoft.TeamFoundationServer.Client -Version 15.112.1
Install-Package Microsoft.TeamFoundationServer.ExtendedClient -Version 15.112.1
Install-Package Microsoft.VisualStudio.Services.Client -Version 15.112.1
Install-Package Microsoft.VisualStudio.Services.InteractiveClient -Version 15.112.1
Install-Package Microsoft.Office.Interop.Excel -Version 15.0.4795.1000
```
Might need to do the steps from the following thread to get Excel working: \
https://learn.microsoft.com/en-us/answers/questions/1496033/microsoft-office-interop-excel-reference-cannot-be

