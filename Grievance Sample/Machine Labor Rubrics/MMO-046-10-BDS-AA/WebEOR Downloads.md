# USPS WebEOR Data Extraction Process

## Overview
This document outlines the step-by-step process for extracting and saving WebEOR (End Of Run) data from the USPS internal system. This process allows users to download operational data in CSV format for analysis and reporting purposes.

## Prerequisites
- Access to USPS internal network
- WebEOR system credentials
- Microsoft Edge browser

## Process Steps

### 1. System Access
1. Launch Microsoft Edge browser
2. Navigate to the USPS Blue homepage
3. Access the WebEOR Welcome Screen

### 2. Site Selection
1. On the WebEOR Login View screen, select your facility from the dropdown menu
2. Click anywhere on the login screen to continue
3. Select "EOR Viewer" from the main menu options

### 3. Date Range Configuration
1. Click the calendar date selector button (...)
2. In the calendar popup:
   - Navigate to the desired month
   - Select the appropriate date
   - Confirm year is correct one year before your desired selection range.
   - Click "OK" to close the calendar

2. Configure date range:
   - For the end date, use the ">>" button to advance through dates (7 day maximum)
   - Verify the date range is displayed correctly in the interface

### 4. Data Retrieval
1. Click the "Refresh Data" button to load all data for the selected date range
2. Once data is loaded, click the "CSV" link in the data view section
3. When the download prompt appears, select "Save as" option

### 5. Saving the CSV File
1. In the Save As dialog:
   - Navigate to the correct folder path:
     `C:\Users\<Your ACE ID>\OneDrive - USPS\Desktop\WebEOR Data`
   - Verify the file name is correct (default: WebEOR-Table-[YYYYMMDD-HHMMSS].csv)
   - Click "Save" to complete the download

### 6. Verification
1. Confirm the file has been saved to the specified location
2. Check that the data covers the intended date range
3. Verify the CSV file contains all required fields and records