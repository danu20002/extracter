# Excel Data Extractor

This Spring Boot application provides REST APIs to extract and process data from Excel files.

## Features

- Extract data from Excel files (.xlsx and .xls formats)
- Process multiple sheets within files
- Perform various data operations and analysis
- RESTful API endpoints for easy integration
- Comprehensive error handling and logging

## Setup

1. **Place Excel Files**: Put your Excel files in the `excel/` folder in the project root
2. **Build the Project**: Run `mvn clean install`
3. **Start the Application**: Run `mvn spring-boot:run` or execute the JAR file

## API Endpoints

### File Management
- `GET /api/excel/health` - Check application health and available files
- `GET /api/excel/files` - List all available Excel files

### Data Extraction
- `GET /api/excel/extract-all` - Extract data from all Excel files
- `GET /api/excel/extract/{fileName}` - Extract data from specific file
- `GET /api/excel/extract/{fileName}/{sheetName}` - Extract data from specific sheet

### Data Operations
- `GET /api/excel/summary` - Get summary of all extracted data
- `GET /api/excel/operations/{operation}` - Perform operation on all data
- `POST /api/excel/operations/{operation}` - Perform operation on provided data

### Available Operations
- `count` - Count total records
- `summary` - Get comprehensive data summary
- `groupbysheet` - Group data by sheet name
- `groupbyfile` - Group data by file name
- `numeric_analysis` - Analyze numeric columns (sum, average, min, max)

## Example Usage

### 1. Check Health
```bash
curl http://localhost:8080/api/excel/health
```

### 2. List Files
```bash
curl http://localhost:8080/api/excel/files
```

### 3. Extract All Data
```bash
curl http://localhost:8080/api/excel/extract-all
```

### 4. Get Data Summary
```bash
curl http://localhost:8080/api/excel/summary
```

### 5. Perform Numeric Analysis
```bash
curl http://localhost:8080/api/excel/operations/numeric_analysis
```

## Response Format

### ExcelProcessingResult
```json
{
  "fileName": "sample.xlsx",
  "success": true,
  "message": "Successfully extracted data",
  "totalSheets": 2,
  "totalRows": 100,
  "sheetNames": ["Sheet1", "Sheet2"],
  "extractedData": [...]
}
```

### ExcelData
```json
{
  "fileName": "sample.xlsx",
  "sheetName": "Sheet1",
  "rowNumber": 2,
  "data": {
    "Name": "John Doe",
    "Age": 30,
    "Salary": 50000
  },
  "extractedAt": "2025-01-15 10:30:00"
}
```

## Error Handling

The application handles various error scenarios:
- Missing Excel files
- Corrupted Excel files
- Invalid sheet names
- File access permissions
- Large file processing

All errors are logged and returned with appropriate HTTP status codes.

## Dependencies

- Spring Boot 3.5.4
- Apache POI 5.2.5 (Excel processing)
- Lombok (reducing boilerplate code)
- Jackson (JSON processing)

## Configuration

The application can be configured through `application.properties`:
- Server port: `server.port=8080`
- Logging levels
- Actuator endpoints

## Notes

- Excel files should be placed in the `excel/` folder
- Both .xlsx (Excel 2007+) and .xls (Excel 97-2003) formats are supported
- The first row is treated as headers
- Empty rows are skipped during processing
- Numeric analysis only works on columns containing numeric data
