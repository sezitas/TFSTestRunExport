# TestCaseExport

Export Test Cases from Team Foundation Server (or Microsoft Test Manager) to Excel.

https://web.archive.org/web/20171108213947/https://tfstestcaseexporttoexcel.codeplex.com/
https://web.archive.org/web/20171108214234/https://tfstestcaseexporttoexcel.codeplex.com/documentation
https://web.archive.org/web/20171108214252/https://tfstestcaseexporttoexcel.codeplex.com/team/view

NOTE: This repository contains changes made to an older project hosted on CodePlex (linked above).

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

