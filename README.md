# FPSO-ETL-pipeline
TLDR; 
ETL (Extract, Transform, Load) pipeline for FPSO actuated valve instrumentation data. Extracts 118 fields across 200+ spec sheets from a DRM-encrypted Excel workbook. Uses Windows COM automation (xlwings) to bypass file-level encryption. Outputs a single consolidated database via openpyxl. 

This tool was made during an internship in Hanwha Ocean, which builds FPSO. 
FPSO stands for Floating Production Storage and Offloading. This platform extracts crude oil from a seabed, which is then processed on the ship. Once processed, the final products of oil and gas are offloaded to other companies' ships. 
In instrumentation & control engineering department, reviewing instruments is an essential task. And there are many many instruments, hence many datasheets. To ease this process, I basically condense many datasheets into a single excel file. 

<img width="1784" height="997" alt="image" src="https://github.com/user-attachments/assets/fd22ba4c-936a-4445-8785-030b254827df" />
