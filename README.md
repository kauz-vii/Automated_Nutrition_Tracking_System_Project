# Automated Nutrition Tracker (Python + Excel Automation)

## Project Overview

**Project Title**: Automated Nutrition Tracking System\
**Level**: Intermediate\
**Tools & Technologies**: Python, openpyxl, Pyperclip, PyInstaller,
Microsoft Excel

This project automates daily nutrition logging by connecting ChatGPT
output directly to an Excel tracker using a background Python
application.

It eliminates manual data entry by:

-   Listening to clipboard activity
-   Detecting structured macro output
-   Automatically updating Excel
-   Applying conditional formatting

The system simulates a real-world personal data pipeline where
user-generated data is captured, processed, and stored automatically.

------------------------------------------------------------------------

## Objectives

1.  Automate daily calorie and macro tracking with zero manual Excel
    entry.
2.  Capture structured output from ChatGPT via clipboard monitoring.
3.  Auto-create daily records based on system date.
4.  Apply business logic for calorie deficit visualization.
5.  Convert the Python script into a background executable for
    real-world usability.

------------------------------------------------------------------------

## Key Features

-   Clipboard listener for real-time data capture
-   Automatic date row creation
-   Macro logging (Calories, Protein, Carbs, Fat)
-   Conditional formatting:
    -   Green → Calorie deficit (\< 2200 kcal)
    -   Red → Calorie surplus (≥ 2200 kcal)
-   Habit tracking columns auto-filled
-   Runs silently in the background
-   Double-click executable (no terminal required)

------------------------------------------------------------------------

## Project Workflow

### 1️⃣ User Interaction

At the end of the day:

FINAL_MACROS: 2135,162,198,68

User copies the line.

### 2️⃣ Clipboard Detection

Python continuously monitors the clipboard and detects the keyword.

### 3️⃣ Automated Excel Update

-   Locate or create today's date
-   Insert macro values
-   Apply formatting logic
-   Save the workbook

------------------------------------------------------------------------

## Business Logic

### Calorie Deficit Rule

  Condition          Format
  ------------------ --------
  Calories \< 2200   Green
  Calories ≥ 2200    Red

### Habit Columns

Always marked as completed:

-   Protein intake
-   10k steps
-   4L+ water

------------------------------------------------------------------------

## Core Implementation

### File Path Handling for EXE Runtime

Ensures the executable always accesses the correct Excel file.

### Auto Create Today's Row

Creates a new row if today's date does not exist.

------------------------------------------------------------------------

## Tools & Technologies

-   Python\
-   openpyxl → Excel automation
-   pyperclip → Clipboard monitoring
-   PyInstaller → Standalone executable
-   Microsoft Excel → Data storage & visualization

------------------------------------------------------------------------

## How to Run the Project

### Development Mode

```python
python auto_macros.py
```

### Production Mode (Executable)

```python
pyinstaller --onefile --noconsole auto_macros.py
```

Then:

1.  Place auto_macros.exe and Macros.xlsx in the same folder
2.  Double-click the EXE
3.  Copy the macro output from ChatGPT

Excel updates automatically.

------------------------------------------------------------------------

## Example Input

FINAL_MACROS: 2050,160,190,65

------------------------------------------------------------------------

## Example Output in Excel

  Date         Calories   Protein   Carbs   Fat   Deficit
  ------------ ---------- --------- ------- ----- ---------
  27-02-2026   2050       160       190     65    Green

------------------------------------------------------------------------

## Key Learnings

-   Event-driven automation using clipboard listeners
-   Excel file manipulation with Python
-   Handling file paths in packaged executables
-   Converting scripts into production-ready background apps
-   Implementing real-world business logic in data workflows

------------------------------------------------------------------------

## Use Cases

-   Personal fitness tracking automation
-   Habit tracking systems
-   Low-code data pipelines
-   Automated logging tools

------------------------------------------------------------------------

## Author -- Kaushik Bhadra

This project is part of my data analytics & automation portfolio and
demonstrates:

-   Process automation using Python
-   Real-world problem solving
-   Excel as a data storage layer
-   Building user-friendly executable tools

------------------------------------------------------------------------

## Portfolio Value

This project showcases:

-   End-to-end automation design
-   Event-driven data ingestion
-   Business rule implementation
-   Python → production deployment
-   Practical, real-life analytics workflow

Daily nutrition logging time reduced from:

2--3 minutes → 2 seconds
