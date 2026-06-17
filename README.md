# FPSO-ETL-pipeline
TLDR; 
ETL pipeline that extracts 118 instrumentation fields from 200+ actuated valve datasheets across a DRM-encrypted Excel workbook, transforming multi-tab P&ID-referenced spec sheets into a single consolidated database via Windows COM automation (xlwings + openpyxl).



This tool was made during an internship in Hanwha Ocean, which builds FPSO. 
FPSO stands for Floating Production Storage and Offloading. This platform extracts crude oil from a seabed, which is then processed on the ship. Once processed, the final products of oil and gas are offloaded to other companies' ships. 
In instrumentation & control engineering department, reviewing instruments is an essential task. And there are many many instruments, hence many datasheets. To ease this process, I basically condense many datasheets into a single excel file. 

<img width="1784" height="997" alt="image" src="https://github.com/user-attachments/assets/fd22ba4c-936a-4445-8785-030b254827df" />
