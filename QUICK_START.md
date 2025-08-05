# Quick Start Guide

## How to Run the Excel Data Extractor

### 1. Start the Application
Open a PowerShell/Command Prompt in the project directory and run:
```bash
.\mvnw.cmd spring-boot:run
```

### 2. Test the Application
Once started (you'll see "Started ExtracterApplication"), open a web browser and go to:
- **http://localhost:8080/** - Test if application is running
- **http://localhost:8080/api/excel/health** - Check Excel processing health

### 3. Add Your Excel Files
- Put your Excel files (.xlsx or .xls) in the `excel/` folder
- The application will automatically detect them

### 4. Extract Data
Use these API endpoints:

#### Get Available Files
```
GET http://localhost:8080/api/excel/files
```

#### Extract All Data
```
GET http://localhost:8080/api/excel/extract-all
```

#### Get Data Summary
```
GET http://localhost:8080/api/excel/summary
```

#### Perform Operations
```
GET http://localhost:8080/api/excel/operations/count
GET http://localhost:8080/api/excel/operations/numeric_analysis
GET http://localhost:8080/api/excel/operations/groupbysheet
```

### 5. Test with Sample Data
I've created a sample CSV file in the excel folder. To test with actual Excel files:
1. Create an Excel file with some data
2. Save it in the `excel/` folder
3. Use the API endpoints above

### Example API Responses

**Health Check:**
```json
{
  "status": "UP",
  "excelFolderExists": true,
  "availableExcelFiles": 1,
  "fileNames": ["your-file.xlsx"]
}
```

**Data Summary:**
```json
{
  "totalRecords": 100,
  "uniqueFiles": 1,
  "uniqueSheets": 2,
  "fileDistribution": {"file1.xlsx": 50, "file2.xlsx": 50},
  "sheetDistribution": {"Sheet1": 75, "Sheet2": 25}
}
```

### Available Operations:
- `count` - Count total records
- `summary` - Get comprehensive summary
- `groupbysheet` - Group data by sheet
- `groupbyfile` - Group data by file
- `numeric_analysis` - Analyze numeric columns

The application runs on **http://localhost:8080**
