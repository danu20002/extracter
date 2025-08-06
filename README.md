# Excel Data Explorer

## Overview
Excel Data Explorer is a powerful web-based tool for extracting, analyzing, and visualizing data from Excel files. The application provides an intuitive interface to upload Excel files, browse their contents, and perform various analytical operations.

## Features

### File Management
- Upload Excel files through an intuitive web interface
- Browse and manage uploaded files
- View detailed file metadata

### Data Extraction
- Extract data from specific sheets or entire workbooks
- View and navigate through sheet data with pagination
- Search within extracted data
- Support for various Excel formats (XLS, XLSX)

### Data Analysis
- Perform statistical analysis on Excel data
- Group data by sheets or files
- Calculate numeric statistics (min, max, average, sum)
- Generate column statistics (data types, unique values, missing values)
- Count records across files and sheets

### Visualization
- Visualize analysis results with interactive charts
- See data distribution across files and sheets
- Bar charts for numeric data analysis
- Pie charts for categorical data distribution

## Technical Implementation

### Backend
- Spring Boot 3.5.x application
- Apache POI for Excel file processing
- Multi-threaded processing for large files
- Memory-mapped file access for improved performance

### Frontend
- Thymeleaf templates with Bootstrap 5
- Interactive data tables using DataTables.js
- Chart visualizations with Chart.js
- Responsive design for desktop and mobile access

## Getting Started

### Prerequisites
- Java 17 or higher
- Maven 3.6 or higher

### Running the Application
1. Clone the repository
2. Navigate to the project directory
3. Run `./mvnw spring-boot:run`
4. Access the application at `http://localhost:8080`

## Usage

1. **Upload Files**: Navigate to the upload page and select Excel files to upload
2. **Browse Files**: View the list of uploaded files and their details
3. **Analyze Data**: Select from various analysis operations:
   - Count records
   - Group by sheet or file
   - Calculate numeric statistics
   - Generate column statistics
4. **View Results**: Explore the analysis results with tables and charts

## Customization

The application can be customized through application.properties:
- File storage location
- Maximum file size
- Processing threads
- Security settings

## License

This project is licensed under the MIT License - see the LICENSE file for details.
