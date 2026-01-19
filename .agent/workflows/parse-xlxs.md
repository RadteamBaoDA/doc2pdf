---
description: Convert xlxs to pdf
---

Implement logic convert or create microsoft office excel to pdf with below requirements:

Goal: Convert or create new microsoft excel to pdf for OCR scan and import to AI knowledge base
Reference Specifications:
1. Add option only convert xlsm, xlsb, xls to xlsx only
2. From xlsx or xlsm file create new pdf file with below requirement
- **STRICLY** pdf file created must have exacly same layout in pdf file as page print(keep same as outline all object, image, theme, color, table border, .......)
- Impelemnt all setting in yml file and logic code in ## Setting area
- If file not have page print, create smart page print with Dimensions is always fit columns, it means ALL COLUMNS VALUE MUST FIT iN PAGE, user can config row dimensions.
- Smart page size, page size must enought for view normal text with size 14, **DO NOT CARE ABOUT PAGE SIZE, ALL CONTENT MUST FIX ON PAGE WIDE IN COLUMN DIMENSIONS**, auto crop page height base on row dinesions config
- Smart metadata, this play metadata of file in header of pdf file as below sample:
   ---
   left: sheet name
   center: rows dimensions index range
   right: file name.
   ---
   eg row dimensions index: begin with 1 as excel, if user config row dimensions is 10 print center metadata as(1 + row dimensions) : 1-10 in page one, 11-20 in page two, xx-yy to final page
## Setting: 
### Implement all below setting when parse to pdf:
1. Sheet name -> seperate config apply each sheet name
2. Orientation: Portrait, landscape
3. Row dimesions: define int num to define how many rows in one pdf page, **SPECICAL VALUE**, null is automatic, 0 is Fit Sheet on One Page
4. Metadata: true|false, true is enable print metadata, otherwise.
Implementation Steps:

Step 1: Define the Schema - Create a Pydantic model that encapsulates all variables from the images.
Step 2: Create the Converter Logic - using OOP to implement this for apply multi create type and office format
Step 3: update CLI and TUI ui for display progress and smart logic using file size and total file remaining to calculate ETA time. TUI display as below
Feature,Implementation Component,Benefit
Progress Bars,rich.progress,Visual feedback for long-running PDF conversions.
Status Spinners,rich.status,Shows the app is working during COM object initialization.
Times, rich.progress, It automatically calculates the Time Elapsed and ETA (Estimated Time of Arrival) based on the average speed of your task.
Tables,rich.table,Displays document metadata or conversion summaries.
Layouts,rich.layout,Creates a split-screen view for logs and progress.
Step 4: Implement Smart calculation works: 
Total Time Running: Tracked by TimeElapsedColumn. It starts the moment the task is added.

ETA Remaining: Tracked by TimeRemainingColumn. If your conversion slows down (e.g., a very large PDF), the ETA will dynamically increase to reflect the real-world speed.
**Logic Comparison:** 
Metric,Calculation Logic,Benefit
Elapsed Time,Tcurrent​−Tstart​,Transparency on how long the process has taken.
ETA (Smart),(Itemsremaining​)×(Average Time per Item),Gives the user a realistic expectation of finish time.

Step5: Implement config to config.yml and apply multi config base on sheet name on excel file