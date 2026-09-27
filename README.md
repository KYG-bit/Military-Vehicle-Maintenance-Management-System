# Military-Vehicle-Maintenance-Management-System

A Python-based vehicle maintenance management system that analyzes vehicle data from Excel and provides an overview of engine oil maintenance status based on vehicle type and mileage.

## Overview

This project was developed to simplify the process of checking vehicle maintenance status from large amounts of vehicle data.

The system reads vehicle information from an Excel file, validates and cleans the data, calculates the mileage driven since the most recent engine oil change, and determines the maintenance status based on vehicle-specific replacement intervals.

The project is being developed incrementally, with plans to extend the system into a GUI-based dashboard and a standalone Windows application.

## Key Features

### Implemented

* Excel-based vehicle data loading
* Required column validation
* Mileage data cleaning and validation
* Calculation of mileage since the most recent engine oil change
* Vehicle-type-specific engine oil replacement intervals
* Automatic maintenance status classification
* Vehicle-type-based maintenance statistics

### Planned

* GUI-based Excel file selection
* Maintenance dashboard
* Visual representation of maintenance statistics
* Vehicle-level maintenance status search
* Standalone Windows executable using PyInstaller

## How It Works

The system determines engine oil maintenance status using the mileage driven since the most recent oil change.

```text
Current Mileage
       -
Mileage at Last Oil Change
       ↓
Mileage Since Oil Change
       ↓
Compare with Vehicle Type Interval
       ↓
Maintenance Status
       ↓
Statistics / Dashboard
```

### Vehicle-Specific Replacement Intervals

| Vehicle Type | Engine Oil Replacement Interval |
| ------------ | ------------------------------: |
| A            |                        5,000 km |
| B            |                        7,000 km |
| C            |                       10,000 km |

### Maintenance Status

| Status    | Description                                                |
| --------- | ---------------------------------------------------------- |
| 정상        | The vehicle is sufficiently below the replacement interval |
| 교체주기 임박   | 500 km or less remain before the replacement interval      |
| 교체 시급     | The replacement interval has been reached or exceeded      |
| 기준 없음     | No replacement interval is configured for the vehicle type |
| 데이터 확인 필요 | Required mileage data is missing or invalid                |

## Tech Stack

* **Language:** Python
* **Data Processing:** Pandas
* **Excel Processing:** openpyxl
* **GUI:** Tkinter *(Planned)*
* **Development Environment:** GitHub Codespaces
* **Version Control:** Git / GitHub
* **Executable Packaging:** PyInstaller *(Planned)*

## Project Structure

```text
Military-Vehicle-Maintenance-Management-System/
│
├── main.py
├── config.py
├── excel_manager.py
├── analyzer.py
│
├── input/
│   └── 차량정보.xlsx
│
├── requirements.txt
└── README.md
```

### File Description

| File               | Description                                                               |
| ------------------ | ------------------------------------------------------------------------- |
| `main.py`          | Controls the overall execution flow of the program                        |
| `config.py`        | Manages configurable vehicle types and maintenance criteria               |
| `excel_manager.py` | Loads, validates, and cleans Excel data                                   |
| `analyzer.py`      | Performs mileage calculations, maintenance classification, and statistics |
| `requirements.txt` | Manages Python package dependencies                                       |
| `README.md`        | Project documentation                                                     |

## Development Process

The project is being developed incrementally, separating data processing, analysis, and user interface responsibilities.

### Phase 1 — Data Processing

* Load vehicle data from Excel
* Validate required columns
* Clean numeric mileage data

### Phase 2 — Maintenance Analysis

* Calculate mileage since the most recent oil change
* Apply vehicle-specific replacement intervals
* Classify maintenance status
* Generate vehicle-type statistics

### Phase 3 — GUI Dashboard

* Implement GUI-based Excel file selection
* Display maintenance statistics
* Visualize vehicle maintenance status

### Phase 4 — Windows Application

* Package the application as a standalone Windows executable
* Allow users to launch the program without directly running Python
* Provide a simple workflow from Excel file selection to maintenance dashboard

## Challenges & Solutions

### Headless Development Environment

The project is being developed in GitHub Codespaces, where a graphical display is not available.

Because Tkinter requires a graphical environment, the Excel file selection process is separated from the core data processing logic.

During development, a predefined Excel file is used:

```text
input/차량정보.xlsx
```

The Windows version will replace this development input method with a GUI-based file selection dialog while keeping the core analysis logic unchanged.

## Future Improvements

The project may be extended to support additional maintenance management features, such as:

* Additional maintenance categories
* Parts management
* Maintenance schedule management
* Vehicle search and filtering
* Dashboard-based reporting
* Data export
* Additional statistical analysis

## Author

Developed as a personal software project using Python and GitHub.
