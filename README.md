# APWU Grievance Documentation & Analysis System

A comprehensive suite of tools for American Postal Workers Union (APWU) representatives to analyze postal operations data, identify contract violations, and generate complete grievance documentation packages.

## System Overview

This integrated system provides APWU representatives with everything needed to:
- **Analyze** postal machine operational data from WebEOR exports
- **Calculate** proper staffing requirements using MMO standards
- **Identify** staffing discrepancies and contract violations
- **Generate** Request for Information (RFI) documents
- **Create** complete Step 1 grievance packages with supporting documentation
- **Document** various types of maintenance and staffing violations

## Table of Contents
1. [Quick Start Guide](#quick-start-guide)
2. [Video Tutorial](#video-tutorial)
3. [System Components](#system-components)
4. [Complete Workflow](#complete-workflow)
5. [Machine Configuration Best Practices](#machine-configuration-best-practices)
6. [Installation & Setup](#installation--setup)
7. [Component Details](#component-details)
8. [Data Sources & Formats](#data-sources--formats)
9. [WebEOR Data Collection](#webeor-data-collection-best-practices)
10. [Troubleshooting](#troubleshooting)
11. [Technical Specifications](#technical-specifications)

---

## Quick Start Guide

### Time Requirements

#### Initial Setup
- Software installation: 15-30 minutes
- WebEOR data collection: 30-60 minutes (52 weekly exports)
- Machine configuration: 2-4 hours (depending on facility size) (As opposed to a month or more manually)
- Report generation and review: 10 minutes

#### Ongoing Maintenance
- Weekly WebEOR updates: 15 minutes
- Quarterly reviews: 1-2 hours

### For Grievance Documentation
1. **Identify the Issue**: Determine the type of violation (staffing, maintenance, etc.)
2. **Generate RFI**: Use `RFI Builder.html` to create initial information request
3. **Analyze Data**: 
   - Use `WebEOR Data Visualizer.html` for operational analysis
   - Use `Machine Checker.ps1` for staffing calculations
4. **Create Grievance**: Use `Grievance Form.html` to compile everything into a complete package

### For Quick Analysis Only
- **Operational Data**: Open `WebEOR Data Visualizer.html` → Import CSV files → Generate report
- **Staffing Analysis**: Run `Machine Checker.ps1` → Import data → Generate staffing report
- **RFI Creation**: Open `RFI Builder.html` → Select template → Fill details → Generate

---
## Video Tutorial

A comprehensive video tutorial (1:29:44) is available covering the entire system setup and usage:
- [YouTube Tutorial Link] - Complete walkthrough from installation to grievance filing
- Chapter timestamps available in video description for easy navigation

---
## System Components

### 1. 📊 WebEOR Data Visualizer (`WebEOR Data Visualizer.html`)
**Purpose**: Analyze postal machine operational data from WebEOR CSV exports

**Key Features**:
- Multi-file CSV import with automatic deduplication
- 6 interactive chart types for operational insights
- Machine performance and utilization analysis
- Filters by date, machine, tour, and site
- Export filtered data and generate analysis reports

**Use Cases**:
- Identify machines with low utilization
- Track operational patterns and trends
- Document equipment performance issues
- Validate staffing calculations with actual usage

### 2. 📝 RFI Builder (`RFI Builder.html`)
**Purpose**: Generate legally formatted Request for Information documents

**Key Features**:
- 6 pre-configured templates for common violations
- Auto-populated legal citations and precedents
- APWU-branded professional formatting
- Save as HTML or print as PDF

**Templates Available**:
- eWHEP staffing package challenges
- General maintenance understaffing
- Custodial understaffing (MS-47)
- Preventive maintenance bypassing
- Maintenance coverage during operations
- Failure to complete staffing packages

### 3. 📋 Grievance Form (`Grievance Form.html`)
**Purpose**: Create complete Step 1 grievance documentation packages

**Key Features**:
- Import Machine Staffing Reports and WebEOR Analysis Reports
- Auto-populate grievance language based on imported data
- Attach supporting documentation within the grievance
- Save complete packages with embedded reports
- Print-ready formatting with proper pagination

**Integration**:
- Accepts reports from both PowerShell tool and WebEOR Visualizer
- Preserves all supporting documentation in a single file
- Maintains chain of evidence for grievance processing

### 4. 🔧 Machine Checker (`Machine Checker.ps1`)
**Purpose**: Calculate proper maintenance staffing based on MMO standards

**Key Features**:
- Import WebEOR data to determine machine usage
- Apply MMO lookup tables for staffing calculations
- Compare current vs. required staffing levels
- Generate detailed discrepancy reports
- Support for all machine types and configurations

**Output**:
- Machine Staffing Reports showing hour discrepancies
- Color-coded surplus/deficit visualization
- MMO-compliant calculations for grievances

### 5. 📂 Supporting Data Files
**CSV Mappings**:
- `Machines_and_Acronyms.csv`: Machine type definitions
- `Machine_PubNum_Mappings.csv`: MMO publication references
- Labor lookup tables per MMO specification
- Current and calculated staffing tables

---

## Complete Workflow

### Standard Grievance Process
```mermaid
graph TD
    A[Identify Contract Violation] --> B{Type of Violation?}
    B -->|Staffing| C[Generate RFI with RFI Builder]
    B -->|Maintenance| C
    B -->|Other| C
    C --> D[Receive Requested Data]
    D --> E[Analyze WebEOR Data]
    E --> F[Run Machine Checker Analysis]
    F --> G[Generate Reports]
    G --> H[Import Reports to Grievance Form]
    H --> I[Complete Grievance Package]
    I --> J[File Step 1 Grievance]
```

### Quick Analysis Workflow
```mermaid
graph LR
    A[WebEOR CSV Files] --> B[WebEOR Data Visualizer]
    B --> C[Operational Analysis Report]
    A --> D[Machine Checker]
    D --> E[Staffing Analysis Report]
    C --> F[Grievance Form]
    E --> F
    F --> G[Complete Documentation]
```
---
### Machine Configuration Best Practices

#### Critical Steps
1. **Always manually select the class code** even if it auto-populates correctly
   - This ensures proper MMO linkage
   - Prevents incorrect labor calculations

2. **Rounding Rules for Tours/Days**:
   - 2.5 or higher → Round up to 3
   - Below 2.5 → Round down to 2
   - 6.9 days → Round to 7
   - Apply consistently across all machines

3. **Special Machine Considerations**:
   - **LAN/MPI**: Shows as "LAN" in system, represents Mail Processing Infrastructure
   - **PSS**: May not have published MMO (check with national)
   - **Renamed Equipment**: If facility renamed machines mid-year, data may be split
---
### Data Validation Checklist

Before generating reports, verify:
- [ ] All machines from eWHEP package are entered
- [ ] Each machine has correct class code selected
- [ ] Staffing data entered for all machines
- [ ] Tours/days match WebEOR analysis (rounded appropriately)
- [ ] All MMO folders contain required lookup tables
---
## Installation & Setup

### Prerequisites
- Windows 10/11 with PowerShell 5.1+
- Modern web browser (Chrome 80+, Firefox 75+, Edge 80+)
- Microsoft Excel or CSV editor
- Administrative privileges for PowerShell execution

### Quick Setup

#### Method 1: Download and Extract
1. **Download the project** from GitHub as ZIP
2. **Extract to a folder** (e.g., `C:\Users\Your Username\Desktop`)
3. **Update paths** in `Machine Checker.ps1`:
   - Open in text editor
   - Find and replace `$global:ProjectRoot = "C:\Users\JR\Desktop\LaborChecklistProd-main` `$global:ProjectRoot = C:\path\to\your\LaborChecklistProd` with your actual path
4. **Verify folder structure**:
   ```
   Your-Folder/
   ├── Machine Checker.ps1
   ├── Grievance Form.html
   ├── WebEOR Data Visualizer.html
   ├── RFI Builder.html
   ├── Staffing Data Visualizer.html
   ├── Mapping CSVs/
   └── Machine Labor Rubrics/
   ```

#### Method 2: Git Clone (Advanced)
```powershell
cd C:\
git clone https://github.com/JRS222/LaborChecklistProd.git APWU-Grievance-System
cd APWU-Grievance-System
# Update paths in Machine Checker.ps1 as above
```

### First Run
1. **For PowerShell Tool**:
   ```powershell
   # Run as Administrator
   Set-ExecutionPolicy Bypass -Scope Process -Force
   .\Machine Checker.ps1
   ```

2. **For HTML Tools**:
   - Simply double-click any `.html` file to open in browser
   - No installation or setup required
   - All tools are self-contained

---

## Component Details

### WebEOR Data Visualizer

#### Data Import Process
1. Click "Select Directory" and choose folder with WebEOR CSV files
2. System automatically processes all CSV files in the directory
3. Duplicates are removed across multiple files
4. Data is validated and parsed for analysis

#### Available Analytics
- **Machine Utilization**: Pie chart of mail volume by machine type
- **Daily Activity**: Line graph of processing volume over time
- **Machine Timeline**: Step chart showing concurrent operations
- **Tour Heatmap**: Hourly activity with throughput overlay
- **Performance Trends**: Top 5 machines tracked over time
- **Tour Distribution**: Workload distribution across tours

#### Report Generation
- Click "Generate Report" for comprehensive HTML analysis
- Includes machine-by-machine statistics
- Shows tours per day and days per week averages
- Identifies underutilized or inefficient machines

### RFI Builder

#### Using Templates
1. **Select Template**: Choose from dropdown based on violation type
2. **Fill Header Fields**: 
   - To/From names and titles
   - References (auto-populated but editable)
   - Particular need (urgency explanation)
3. **Add Facility Info**: Enter installation name
4. **Generate**: Creates formatted RFI with all legal citations

#### Template Customization
- All generated text is editable before saving
- Legal citations are automatically included
- Penalty information from recent cases included
- Professional APWU formatting applied

### Grievance Form

#### Import Capabilities
1. **Machine Staffing Report**:
   - Validates report structure
   - Extracts discrepancy data
   - Auto-populates Background with violation details
   - Auto-populates Corrective Action with remedies

2. **WebEOR Analysis Report**:
   - Serves as supporting documentation
   - Validates staffing calculations
   - Attached to grievance package
   - No text generation (validation only)

#### Form Sections
- **Header Table**: Standard grievance information fields
- **Background**: Auto-populated violation description
- **Corrective Action**: Auto-populated requested remedies  
- **Management Response**: Space for Step 1 response
- **Attachments**: Embedded supporting reports

#### Save Options
- **Save Form**: Downloads complete HTML with all data and attachments
- **Print as PDF**: Formats for printing with proper pagination
- **Auto-save**: Browser localStorage preserves work in progress

### Machine Checker (PowerShell)

#### Machine Configuration
1. **Add Machines**: Manual entry or CSV import
2. **Set Parameters**: 
   - Class code selection (determines MMO)
   - Operational days and tours
   - Machine-specific settings
3. **View Details**: 
   - Labor lookup tables
   - Current staffing
   - Calculated requirements

#### Staffing Calculations
- **Lookup Table Matching**: Finds best match for machine parameters
- **Hour Calculations**: Determines annual maintenance hours needed
- **Skill Distribution**: Allocates hours across MM7, MPE9, ET10
- **Discrepancy Analysis**: Compares required vs. actual staffing

#### Report Features
- **Color Coding**: Green (surplus), Red (deficit)
- **MMO Grouping**: Organized by maintenance publication
- **Summary Totals**: Overall staffing surplus/deficit
- **Machine Details**: Complete configuration for each machine

---

## Data Sources & Formats

### Labor Lookup Tables
The system uses flattened lookup tables extracted from MMOs:
- Located in `Machine Labor Rubrics/[MMO-XXX-XX]/`
- Format: `*-Labor-Lookup.csv`
- Contains operational maintenance hours calculations

### Creating Missing Lookup Tables
If an MMO is missing:
1. Obtain PDF from management or APWU National
2. Extract labor tables manually
3. Create CSV with required columns
4. Contact project maintainer for assistance

### WebEOR CSV Format
```csv
Site,MType,MNo,Op No.,Sort Program,Tour,Run#,Start,End,Fed,MODS,DOIS
"02301-9997","ATU","2","207000","ATU","2","3","04/21/24 07:00","04/21/24 15:00","765","S 04/21/24","NS"
```

### Machine Mappings CSV
```csv
FullName,Acronym
"AUTOMATED FACER CANCELER SYSTEM",AFCS
"ADVANCED FACER CANCELER SYSTEM",AFCS100
```

### MMO Mappings CSV
```csv
Acronym,Class Code,Pub Num
AFCS,AA,MMO-058-21
AFCS100,AA,MMO-077-20
```

### Labor Lookup Table
```csv
Operational Days,Tours/Day,Operational Maintenance (hrs/yr),Total (hrs/yr)
5,1,251.33,1425.58
6,1,301.60,1672.90
7,1,351.87,1920.23
```

---
### WebEOR Data Collection Best Practices

#### CRITICAL: Filter Settings
- **Always select "Without Maintenance Runs"** when exporting WebEOR data
- Using "All Runs" will incorrectly inflate tour calculations
- Maintenance test runs should be excluded from operational data

#### Data Collection Process
1. Navigate to WebEOR in Microsoft Edge (internal network only)
2. Select your facility
3. Set date range to 7 days (maximum allowed)
4. **Filter: Without Maintenance Runs** ← Critical setting
5. Export as CSV
6. Repeat for 52 weeks (one year of data)
7. Save all files in a single "WebEOR Data" folder

**Note**: Expect some data overlap when collecting weekly files. The system automatically removes duplicates during import.

### Excel Integration (Optional but Recommended)

The system works without Excel, but Excel provides helpful features:

1. **Alphabetizing Machine List**:
   - Open exported CSV in Excel
   - Press Ctrl+T (create table)
   - Sort A-Z by machine column
   - Save (Ctrl+S) - keeps CSV format
   - Reimport to PowerShell for organized review

2. **Alternatives to Excel**:
   - OpenOffice Calc (free)
   - Google Sheets
   - Any CSV editor
---
## Troubleshooting

### Common Issues

#### PowerShell Execution Blocked
```powershell
# Solution: Run as Administrator
Set-ExecutionPolicy Bypass -Scope Process -Force
# Or unblock the file
Unblock-File ".\Machine Checker.ps1"
```

#### HTML Tools Won't Open
- **Cause**: Browser blocking local file JavaScript
- **Solution**: Use Chrome/Edge and allow local file access
- **Alternative**: Host files on local web server

#### Import Failures
- **Machine Staffing Report**: Ensure filename contains "machine_staffing_report"
- **WebEOR Report**: Ensure filename contains "webeor_analysis_report"
- **CSV Files**: Check for proper headers and date formats

#### Large Dataset Performance
- **WebEOR Visualizer**: Limit date range for datasets >5,000 records
- **Machine Checker**: Import may take 2-5 minutes for >1,000 machines
- **Browser Memory**: Close other tabs when processing large files

### Debug Mode

#### Browser Console (F12)
- Check for JavaScript errors
- Monitor file loading progress
- Verify data parsing results

#### PowerShell Verbose Mode
```powershell
$VerbosePreference = "Continue"
.\Machine Checker.ps1
```
#### "Not Responding" During Import
- Normal for large datasets (>10,000 records)
- Wait up to 5 minutes for processing
- Check PowerShell window for progress updates

#### Machine Not in WebEOR Data
Some machines don't report to WebEOR:
- Container unloaders
- BDS (Barcode Distribution System)
- DPRC
- Manual operations

For these machines:
1. Add manually using dropdown
2. Use management's stated tours/days from eWHEP package
3. Note in grievance that operational data unavailable

#### Duplicate Machine Entries
If machines appear renamed (e.g., AFSM 100 4 → AFSM100 1):
- Facility renamed equipment during data period
- May need to manually consolidate data
- Document both names in grievance
---

## Technical Specifications

### System Requirements
- **OS**: Windows 10/11, Server 2016+
- **PowerShell**: Version 5.1 or higher
- **Browser**: Chrome 80+, Firefox 75+, Edge 80+, Safari 14+
- **RAM**: 4GB minimum, 8GB recommended
- **Storage**: 500MB for application and data

### Browser Technologies Used
- **File System Access API**: For directory selection
- **localStorage**: For auto-save functionality
- **Chart.js 3.9.1**: For data visualization
- **DOMParser**: For HTML report parsing
- **Blob API**: For file generation

### Security Considerations
- **PowerShell**: Requires execution policy adjustment
- **File Access**: Browser requires permission for local files
- **Data Privacy**: All processing is local, no external transmission
- **Report Sharing**: Generated HTML files are self-contained

### Performance Benchmarks
- **CSV Import**: ~1,000 records/second
- **Chart Rendering**: <2 seconds for 5,000 data points
- **Report Generation**: <5 seconds for 100 machines
- **File Save**: Instant for reports <10MB

---

## Version History

### Current Version: 1.0
- **Added**: RFI Builder for information requests
- **Enhanced**: Grievance Form with dual report import
- **Improved**: WebEOR Visualizer with 6 chart types
- **Updated**: Machine Checker with better import handling

### Future Enhancements
- Cloud storage integration
- Multi-facility report consolidation
- Automated MMO update checking
- Mobile-responsive design improvements

---

## Support & Resources

### Documentation
- This README file
- Inline help in each application
- MMO reference PDFs in Machine Labor Rubrics folders

### Reporting Issues
When reporting problems, include:
1. Which tool has the issue
2. Steps to reproduce
3. Error messages (screenshot or text)
4. Sample data (remove sensitive information)
5. Browser/PowerShell version

### Contributing
This system is maintained by APWU representatives. To contribute:
1. Test changes thoroughly
2. Document new features
3. Maintain backwards compatibility
4. Follow existing code style

---

### Planned Improvements
- Python-based backend
- Web interface for easier use
- Support for building maintenance (MS-1)
- Support for custodial (MS-47)
- Automated WebEOR data fetching

### Contact for Support
- GitHub Issues: [Link to issues page](https://github.com/JRS222/LaborChecklistProd/issues)
- Email: joseph.r.shavv@gmail.com
- APWU Maintenance Craft Facebook Group

*APWU Grievance Documentation & Analysis System - Empowering union representatives with data-driven grievance tools*
