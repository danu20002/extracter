package com.jnj.extracter.serviceImpl;

import com.jnj.extracter.entity.ExcelData;
import com.jnj.extracter.entity.ExcelProcessingResult;
import com.jnj.extracter.service.ExcelService;
import com.jnj.extracter.util.ExcelParsingUtils;
import lombok.extern.slf4j.Slf4j;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.commons.io.FilenameUtils;

@Slf4j
@Service
public class ExcelServiceImpl implements ExcelService {

    private static final String EXCEL_FOLDER_PATH = "excel";
    private static final String TEMP_FOLDER_PATH = "excel/temp";
    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    @Override
    public List<ExcelProcessingResult> extractAllExcelFiles() {
        List<File> excelFiles = getExcelFiles();
        List<ExcelProcessingResult> results = new ArrayList<>();
        
        for (File file : excelFiles) {
            try {
                ExcelProcessingResult result = extractExcelFile(file);
                results.add(result);
                log.info("Successfully processed file: {}", file.getName());
            } catch (Exception e) {
                log.error("Error processing file: {}", file.getName(), e);
                ExcelProcessingResult errorResult = new ExcelProcessingResult();
                errorResult.setFileName(file.getName());
                errorResult.setSuccess(false);
                errorResult.setMessage("Error: " + e.getMessage());
                results.add(errorResult);
            }
        }
        
        return results;
    }

    @Override
    public ExcelProcessingResult extractExcelFile(File file) {
        ExcelProcessingResult result = new ExcelProcessingResult();
        result.setFileName(file.getName());
        
        // Create temp directory if it doesn't exist
        File tempDir = new File(TEMP_FOLDER_PATH);
        if (!tempDir.exists()) {
            tempDir.mkdirs();
        }
        
        // Handle XLSB files separately
        File fileToProcess = file;
        
        if (file.getName().toLowerCase().endsWith(".xlsb")) {
            log.info("Detected XLSB file: {}. Will attempt to handle it specially.", file.getName());
            result.setMessage("Warning: XLSB files have limited support. Some data may not be extracted correctly.");
            
            try {
                // We cannot directly process XLSB, but we'll try with WorkbookFactory as a fallback
                fileToProcess = file;
            } catch (Exception e) {
                log.error("XLSB processing error: {}", e.getMessage());
                result.setSuccess(false);
                result.setMessage("Error: This XLSB file cannot be processed. Please convert it to .xlsx format.");
                return result;
            }
        }
        
        try (FileInputStream fis = new FileInputStream(fileToProcess)) {
            // Set additional security settings for processing potentially problematic files
            org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(0.0001); // More permissive ratio
            org.apache.poi.openxml4j.util.ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024); // 100MB max entry size
            
            Workbook workbook = createWorkbook(fileToProcess, fis);
            
            List<String> sheetNames = new ArrayList<>();
            List<ExcelData> allData = new ArrayList<>();
            int totalRows = 0;
            
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                sheetNames.add(sheetName);
                
                List<ExcelData> sheetData = extractDataFromSheet(sheet, file.getName());
                allData.addAll(sheetData);
                totalRows += sheetData.size();
            }
            
            result.setSuccess(true);
            result.setMessage("Successfully extracted data");
            result.setTotalSheets(workbook.getNumberOfSheets());
            result.setTotalRows(totalRows);
            result.setSheetNames(sheetNames);
            result.setExtractedData(allData);
            
            workbook.close();
            
        } catch (Exception e) {
            log.error("Error extracting data from file: {}", file.getName(), e);
            result.setSuccess(false);
            
            // Provide more specific error messages
            String errorMessage = e.getMessage();
            String errorString = e.toString();
            
            if (file.getName().toLowerCase().endsWith(".xlsb") && errorString.contains("XLSBUnsupported")) {
                errorMessage = "XLSB file processing error. Make sure you have the correct dependencies for XLSB support.";
            } else if (errorString.contains("Zip bomb detected") || errorString.contains("ZipSecureFile") || errorString.contains("ratio")) {
                // Handle zip bomb detection errors
                errorMessage = "Excel file security check failed. File contains unusually compressed data. " +
                               "The application has been configured to bypass this check. Please try again.";
                
                // Try to bypass the check for next time
                ZipSecureFile.setMinInflateRatio(0.0001); // Very permissive ratio
                ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024); // 100MB max entry size
            } else if (errorString.contains("Invalid byte") && errorString.contains("UTF-8 sequence")) {
                // Handle UTF-8 encoding issues, often found in pivot tables with non-standard characters
                errorMessage = "The Excel file contains data with invalid character encoding, likely in pivot tables. " +
                              "Attempting to process the file with alternative methods.";
                
                // Try to extract data using alternative methods
                try {
                    log.info("Attempting to extract data from file with encoding issues: {}", file.getName());
                    
                    // Get sheet names using the safe method
                    List<String> sheetNames = ExcelParsingUtils.extractSheetNamesFromProblemFile(file);
                    result.setSheetNames(sheetNames);
                    result.setSuccess(true);
                    result.setMessage("Successfully extracted sheet names, but couldn't extract full data due to encoding issues in pivot tables");
                    result.setTotalSheets(sheetNames.size());
                    // We couldn't extract the data so set totalRows to 0
                    result.setTotalRows(0);
                    return result;
                    
                } catch (Exception ex) {
                    log.error("Alternative extraction failed: {}", ex.getMessage());
                    errorMessage = "File contains corrupted data that cannot be processed: " + ex.getMessage();
                }
            } else if (e instanceof IOException) {
                errorMessage = "I/O error while reading the file: " + e.getMessage();
            }
            
            result.setMessage("Error: " + errorMessage);
            // Log the stack trace for debugging
            StringBuffer stackTrace = new StringBuffer();
            for (StackTraceElement element : e.getStackTrace()) {
                stackTrace.append(element.toString()).append("\n");
            }
            log.debug("Stack trace: {}", stackTrace.toString());
        }
        
        return result;
    }

    @Override
    public List<ExcelData> extractSheetData(File file, String sheetName) {
        List<ExcelData> data = new ArrayList<>();
        
        try (FileInputStream fis = new FileInputStream(file)) {
            Workbook workbook = createWorkbook(file, fis);
            Sheet sheet = workbook.getSheet(sheetName);
            
            if (sheet != null) {
                data = extractDataFromSheet(sheet, file.getName());
            } else {
                log.warn("Sheet '{}' not found in file '{}'", sheetName, file.getName());
            }
            
            workbook.close();
            
        } catch (Exception e) {
            log.error("Error extracting data from sheet '{}' in file '{}'", sheetName, file.getName(), e);
        }
        
        return data;
    }

    @Override
    public List<File> getExcelFiles() {
        File excelDir = new File(EXCEL_FOLDER_PATH);
        List<File> excelFiles = new ArrayList<>();
        
        if (excelDir.exists() && excelDir.isDirectory()) {
            File[] files = excelDir.listFiles((dir, name) -> {
                String lowercase = name.toLowerCase();
                
                // Only process files in the root excel directory, not in temp folders
                if (dir.getName().equals("temp")) {
                    return false;
                }
                
                return lowercase.endsWith(".xlsx") || 
                       lowercase.endsWith(".xls") ||
                       lowercase.endsWith(".xlsb");
            });
            
            if (files != null) {
                excelFiles.addAll(Arrays.asList(files));
                for (File file : files) {
                    if (file.getName().toLowerCase().endsWith(".xlsb")) {
                        log.info("Found XLSB file: {}. Note: XLSB format has limited support.", file.getName());
                    }
                }
            }
        } else {
            log.warn("Excel directory '{}' does not exist", EXCEL_FOLDER_PATH);
        }
        
        return excelFiles;
    }

    @Override
    public Map<String, Object> performDataOperations(List<ExcelData> data, String operation) {
        Map<String, Object> result = new HashMap<>();
        
        switch (operation.toLowerCase()) {
            case "count":
                result.put("totalRecords", data.size());
                result.put("operation", "count");
                break;
                
            case "summary":
                result = getDataSummary(data);
                break;
                
            case "groupbysheet":
                Map<String, Long> groupedBySheet = data.stream()
                    .collect(Collectors.groupingBy(ExcelData::getSheetName, Collectors.counting()));
                result.put("groupedBySheet", groupedBySheet);
                result.put("operation", "groupBySheet");
                break;
                
            case "groupbyfile":
                Map<String, Long> groupedByFile = data.stream()
                    .collect(Collectors.groupingBy(ExcelData::getFileName, Collectors.counting()));
                result.put("groupedByFile", groupedByFile);
                result.put("operation", "groupByFile");
                break;
                
            case "numeric_analysis":
                result = performNumericAnalysis(data);
                break;
                
            default:
                result.put("error", "Unknown operation: " + operation);
                result.put("availableOperations", Arrays.asList("count", "summary", "groupBySheet", "groupByFile", "numeric_analysis"));
        }
        
        return result;
    }

    @Override
    public Map<String, Object> getDataSummary(List<ExcelData> data) {
        Map<String, Object> summary = new HashMap<>();
        
        summary.put("totalRecords", data.size());
        summary.put("uniqueFiles", data.stream().map(ExcelData::getFileName).distinct().count());
        summary.put("uniqueSheets", data.stream().map(ExcelData::getSheetName).distinct().count());
        
        // File distribution
        Map<String, Long> fileDistribution = data.stream()
            .collect(Collectors.groupingBy(ExcelData::getFileName, Collectors.counting()));
        summary.put("fileDistribution", fileDistribution);
        
        // Sheet distribution
        Map<String, Long> sheetDistribution = data.stream()
            .collect(Collectors.groupingBy(ExcelData::getSheetName, Collectors.counting()));
        summary.put("sheetDistribution", sheetDistribution);
        
        return summary;
    }

    private Workbook createWorkbook(File file, FileInputStream fis) throws IOException {
        try {
            // Set zip parameters to bypass zip bomb detection for this specific operation
            // This allows processing of Excel files with unusual compression ratios
            org.apache.poi.openxml4j.util.ZipSecureFile.setMinInflateRatio(0.0001); // More permissive ratio
            org.apache.poi.openxml4j.util.ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024); // 100MB max entry size
            
            String fileName = file.getName().toLowerCase();
            if (fileName.endsWith(".xlsb")) {
                log.warn("XLSB file detected: {}. Attempting direct processing, but this format has limited support.", file.getName());
                
                // Try the more reliable approach for .xlsb files
                try {
                    // Use WorkbookFactory as a first attempt
                    return WorkbookFactory.create(fis);
                } catch (Exception e) {
                    log.error("Standard processing of XLSB failed: {}", e.getMessage());
                    
                    // Fall back to direct OPC approach
                    fis.close();
                    try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
                        XSSFWorkbook workbook = new XSSFWorkbook(pkg);
                        return workbook;
                    } catch (Exception ex) {
                        log.error("Direct OPC processing failed: {}", ex.getMessage());
                        throw new IOException("This XLSB file format cannot be processed. Please convert it to XLSX format.", ex);
                    }
                }
            } else {
                // For other Excel files, try using custom loading to handle corrupted pivot tables
                try {
                    // First try the standard approach
                    return WorkbookFactory.create(fis);
                } catch (Exception e) {
                    // If the error is related to invalid UTF-8 in pivot cache records, try a custom approach
                    if (e.toString().contains("Invalid byte") && e.toString().contains("UTF-8 sequence")) {
                        log.warn("Encountered UTF-8 encoding issue in file: {}. Attempting alternative loading method.", file.getName());
                        
                        // Close the current stream and open a new one
                        fis.close();
                        
                        try {
                            // Try the ExcelParsingUtils approach to extract data while avoiding pivot tables
                            XSSFWorkbook workbook = ExcelParsingUtils.createPartialWorkbook(file);
                            if (workbook != null) {
                                log.info("Successfully loaded file using partial workbook approach: {}", file.getName());
                                return workbook;
                            }
                            
                            log.warn("Partial workbook approach failed, trying direct stream approach");
                            
                            // Use event-based approach instead of dom-based to skip problematic parts
                            try (FileInputStream newFis = new FileInputStream(file)) {
                                // Set custom properties for parsing to handle corrupted files
                                System.setProperty("org.apache.poi.xssf.parsemode", "tolerant");
                                
                                // For .xlsx files
                                if (fileName.endsWith(".xlsx")) {
                                    // Try with strict OOXML disabled
                                    System.setProperty("org.apache.poi.ooxml.strict", "false");
                                    return new XSSFWorkbook(newFis);
                                } 
                                // For .xls files
                                else if (fileName.endsWith(".xls")) {
                                    return new HSSFWorkbook(newFis);
                                } else {
                                    throw new IOException("Unsupported file format after attempting recovery: " + fileName);
                                }
                            }
                        } catch (Exception ex) {
                            log.error("Alternative loading methods failed: {}", ex.getMessage());
                            throw new IOException("File contains corrupted data that cannot be processed. The file may have invalid character encoding or corrupted pivot tables.", ex);
                        }
                    } else {
                        // For other errors, just propagate
                        throw e;
                    }
                }
            }
        } catch (Exception e) {
            log.error("Error creating workbook for file: {}", file.getName(), e);
            throw new IOException("Failed to create workbook: " + e.getMessage(), e);
        }
    }

    private List<ExcelData> extractDataFromSheet(Sheet sheet, String fileName) {
        List<ExcelData> sheetData = new ArrayList<>();
        
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return sheetData;
        }
        
        // Get header row and determine all columns dynamically
        Row headerRow = sheet.getRow(sheet.getFirstRowNum());
        List<String> headers = new ArrayList<>();
        int maxColumns = 0;
        
        // First pass: determine the maximum number of columns across all rows
        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                maxColumns = Math.max(maxColumns, row.getLastCellNum());
            }
        }
        
        // Extract headers from the first row, extending to maxColumns
        if (headerRow != null) {
            for (int cellIndex = 0; cellIndex < maxColumns; cellIndex++) {
                Cell cell = headerRow.getCell(cellIndex);
                String headerValue = getCellValueAsString(cell);
                
                // If header is empty, generate a default name
                if (headerValue == null || headerValue.trim().isEmpty()) {
                    headerValue = "Column_" + (cellIndex + 1);
                }
                headers.add(headerValue);
            }
        } else {
            // No header row found, generate default headers
            for (int i = 0; i < maxColumns; i++) {
                headers.add("Column_" + (i + 1));
            }
        }
        
        // Process data rows (skip header row)
        for (int rowIndex = sheet.getFirstRowNum() + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            
            Map<String, Object> rowData = new HashMap<>();
            boolean hasData = false;
            
            // Extract data for all columns
            for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                Cell cell = row.getCell(cellIndex);
                String header = headers.get(cellIndex);
                Object cellValue = getCellValue(cell);
                
                if (cellValue != null && !cellValue.toString().trim().isEmpty()) {
                    hasData = true;
                }
                
                rowData.put(header, cellValue);
            }
            
            // Only add row if it contains some data
            if (hasData) {
                ExcelData excelData = new ExcelData();
                excelData.setFileName(fileName);
                excelData.setSheetName(sheet.getSheetName());
                excelData.setRowNumber(rowIndex + 1);
                excelData.setData(rowData);
                excelData.setExtractedAt(LocalDateTime.now().format(FORMATTER));
                
                sheetData.add(excelData);
            }
        }
        
        log.info("Extracted {} rows from sheet '{}' with {} columns", 
                sheetData.size(), sheet.getSheetName(), headers.size());
        
        return sheetData;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    return cell.getNumericCellValue();
                } catch (Exception e) {
                    return cell.getStringCellValue();
                }
            default:
                return null;
        }
    }

    private String getCellValueAsString(Cell cell) {
        Object value = getCellValue(cell);
        return value != null ? value.toString() : "";
    }

    private Map<String, Object> performNumericAnalysis(List<ExcelData> data) {
        Map<String, Object> analysis = new HashMap<>();
        Map<String, List<Double>> numericColumns = new HashMap<>();
        
        // Collect numeric data by column
        for (ExcelData row : data) {
            for (Map.Entry<String, Object> entry : row.getData().entrySet()) {
                String column = entry.getKey();
                Object value = entry.getValue();
                
                if (value instanceof Number) {
                    numericColumns.computeIfAbsent(column, k -> new ArrayList<>())
                               .add(((Number) value).doubleValue());
                }
            }
        }
        
        // Perform analysis for each numeric column
        Map<String, Map<String, Double>> columnAnalysis = new HashMap<>();
        for (Map.Entry<String, List<Double>> entry : numericColumns.entrySet()) {
            String column = entry.getKey();
            List<Double> values = entry.getValue();
            
            if (!values.isEmpty()) {
                Map<String, Double> stats = new HashMap<>();
                stats.put("count", (double) values.size());
                stats.put("sum", values.stream().mapToDouble(Double::doubleValue).sum());
                stats.put("average", values.stream().mapToDouble(Double::doubleValue).average().orElse(0.0));
                stats.put("min", values.stream().mapToDouble(Double::doubleValue).min().orElse(0.0));
                stats.put("max", values.stream().mapToDouble(Double::doubleValue).max().orElse(0.0));
                
                columnAnalysis.put(column, stats);
            }
        }
        
        analysis.put("numericColumnAnalysis", columnAnalysis);
        analysis.put("operation", "numeric_analysis");
        
        return analysis;
    }
}
