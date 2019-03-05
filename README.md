# QAC Plugin - DAPA Metrics Report Generation

This is PRQA Framework Plugin for generating Report based on DAPA Source Code Metrics criteria.

## Overview

You can simply generate DAPA Metrics Report by clicking the menu [Report]-[Generate Report for Project: "~~~" ] in PRQA Framework.

![](assets/markdown-img-paste-20190306001010484.png)

In the report, you can see what functions exceeded the criteria of DAPA Source Code Metrics.
The value of metrics is highlighted by red if the function exceeded the limitation of the Metrics.

![](assets/markdown-img-paste-20190306000941901.png)

## Usage

You have to copy these to 'c:\PRQA\PRQA-Framework-2.x.x\report_plugins' directory.
  * 'DAPA_Metrics_Report.py' file
  *  python library 'openpyxl'

![](assets/markdown-img-paste-20190306002631894.png)

You can download 'openpyxl' library at here: https://pypi.org/project/openpyxl/#files

I had used openpyxl 2.6.1 version.
Download 'openpyxl-2.6.1.tar.gz' file and decompress it. Then, move the folder named 'openpyxl' to there.
