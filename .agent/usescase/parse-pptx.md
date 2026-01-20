---
description: Convert doc to pptx
---

Implement logic convert microsoft office powerpoint to pdf with below requirements:

Goal: Convert microsoft word to pdf for OCR scan and import to AI knowledge base
Reference Specifications:

## Setting: 
### Implement all below setting when parse to pdf:
1. Scope: Full page slide, slide range.
2 Color: color, grayscale, pure black and white

Implementation Steps:

Step 1: Define the Schema - Create a Pydantic model that encapsulates all variables from the images.
Step 2: Create the Converter Logic - using OOP to implement this for apply multi converter type and office format
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

Step5: Implement config to config.yml and apply multi config base on folder name like words