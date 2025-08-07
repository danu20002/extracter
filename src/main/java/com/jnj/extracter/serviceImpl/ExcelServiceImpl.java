package com.jnj.extracter.serviceImpl;

import com.jnj.extracter.config.ExcelProcessingConfig;
import com.jnj.extracter.entity.ExcelData;
import com.jnj.extracter.entity.ExcelProcessingResult;
import com.jnj.extracter.service.ExcelService;
import com.jnj.extracter.service.MetricsService;
import com.jnj.extracter.util.ExcelParsingUtils;
import com.jnj.extracter.util.MemoryMappedFileHandler;
import com.jnj.extracter.util.ProtoConverter;
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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.stream.Collectors;
import org.apache.commons.io.FilenameUtils;

@Slf4j
@Service
public class ExcelServiceImpl implements ExcelService {

    private final ExcelProcessingConfig config;
    private final MemoryMappedFileHandler memoryMapper;
    private final MetricsService metricsService;
    private final ProtoConverter protoConverter;
    private final ExecutorService executorService;
    
    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private static final String TEMP_FOLDER_PATH = "excel/temp";
    
    @Autowired
    public ExcelServiceImpl(ExcelProcessingConfig config, 
                           MemoryMappedFileHandler memoryMapper, 
                           MetricsService metricsService,
                           ProtoConverter protoConverter) {
        this.config = config;
        this.memoryMapper = memoryMapper;
        this.metricsService = metricsService;
        this.protoConverter = protoConverter;
        this.executorService = Executors.newFixedThreadPool(config.getThreadPoolSize());
        
        // Initialize POI settings globally
        ZipSecureFile.setMinInflateRatio(0.0001);
        ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024);
        System.setProperty("org.apache.poi.xssf.parsemode", "tolerant");
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        
        log.info("Excel Service initialized with parallel processing={}, threadPoolSize={}, useMemoryMapped={}",
                config.isParallelProcessing(), config.getThreadPoolSize(), config.isUseMemoryMapped());
    }

    @Override
    public List<ExcelProcessingResult> extractAllExcelFiles() {
        Instant startTime = Instant.now();
        List<File> excelFiles = getExcelFiles();
        List<ExcelProcessingResult> results = new ArrayList<>();
        
        if (config.isParallelProcessing() && excelFiles.size() > 1) {
            log.info("Using parallel processing for {} Excel files", excelFiles.size());
            
            try {
                // Process files in parallel using CompletableFuture
                List<CompletableFuture<ExcelProcessingResult>> futures = excelFiles.stream()
                    .map(file -> CompletableFuture.supplyAsync(() -> {
                        try {
                            ExcelProcessingResult result = extractExcelFile(file);
                            log.info("Successfully processed file: {}", file.getName());
                            return result;
                        } catch (Exception e) {
                            log.error("Error processing file: {}", file.getName(), e);
                            ExcelProcessingResult errorResult = new ExcelProcessingResult();
                            errorResult.setFileName(file.getName());
                            errorResult.setSuccess(false);
                            errorResult.setMessage("Error: " + e.getMessage());
                            metricsService.recordProcessingError(e.getClass().getSimpleName());
                            return errorResult;
                        }
                    }, executorService))
                    .collect(Collectors.toList());
                
                // Wait for all futures to complete and collect results
                CompletableFuture<Void> allFutures = CompletableFuture.allOf(
                    futures.toArray(new CompletableFuture[0])
                );
                
                // Get all results
                results = allFutures.thenApply(v -> 
                    futures.stream()
                        .map(CompletableFuture::join)
                        .collect(Collectors.toList())
                ).get();
                
            } catch (Exception e) {
                log.error("Error in parallel processing: {}", e.getMessage(), e);
                // Fall back to sequential processing
                return extractAllFilesSequentially(excelFiles);
            }
        } else {
            // Sequential processing
            results = extractAllFilesSequentially(excelFiles);
        }
        
        Instant endTime = Instant.now();
        long elapsedTime = Duration.between(startTime, endTime).toMillis();
        log.info("Processed {} Excel files in {} ms", excelFiles.size(), elapsedTime);
        metricsService.recordMemoryUsage();
        
        return results;
    }
    
    /**
     * Process all Excel files sequentially.
     * 
     * @param excelFiles List of Excel files to process
     * @return List of processing results
     */
    private List<ExcelProcessingResult> extractAllFilesSequentially(List<File> excelFiles) {
        List<ExcelProcessingResult> results = new ArrayList<>();
        
        for (File file : excelFiles) {
            Instant fileStartTime = Instant.now();
            try {
                ExcelProcessingResult result = extractExcelFile(file);
                results.add(result);
                log.info("Successfully processed file: {}", file.getName());
                
                // Record metrics
                String extension = FilenameUtils.getExtension(file.getName()).toLowerCase();
                metricsService.recordFileProcessed(extension);
                if (result.getTotalRows() > 0) {
                    metricsService.recordRowsProcessed(result.getTotalRows());
                }
                
            } catch (Exception e) {
                log.error("Error processing file: {}", file.getName(), e);
                ExcelProcessingResult errorResult = new ExcelProcessingResult();
                errorResult.setFileName(file.getName());
                errorResult.setSuccess(false);
                errorResult.setMessage("Error: " + e.getMessage());
                results.add(errorResult);
                metricsService.recordProcessingError(e.getClass().getSimpleName());
            } finally {
                Instant fileEndTime = Instant.now();
                long fileElapsedTime = Duration.between(fileStartTime, fileEndTime).toMillis();
                metricsService.recordFileProcessingTime(file.getName(), fileElapsedTime);
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
    public List<String> getSheetNames(File file) {
        if (!file.exists() || !file.isFile()) {
            log.warn("Excel file does not exist: {}", file.getAbsolutePath());
            return Collections.emptyList();
        }
        
        String ext = FilenameUtils.getExtension(file.getName()).toLowerCase();
        if (!ext.equals("xlsx") && !ext.equals("xls") && !ext.equals("xlsb")) {
            log.warn("Not an Excel file: {}", file.getName());
            return Collections.emptyList();
        }
        
        try (Workbook workbook = WorkbookFactory.create(file)) {
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
            return sheetNames;
        } catch (Exception e) {
            log.error("Error getting sheet names from file: {}", file.getName(), e);
            return Collections.emptyList();
        }
    }

    @Override
    public List<File> getExcelFiles() {
        File excelDir = new File(config.getExcelFolderPath());
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
                       lowercase.endsWith(".xlsb") ||
                       lowercase.endsWith(".csv"); // Added CSV support
            });
            
            if (files != null) {
                log.info("Found {} Excel files in directory: {}", files.length, config.getExcelFolderPath());
                excelFiles.addAll(Arrays.asList(files));
                
                // Log details about each file
                for (File file : files) {
                    String extension = FilenameUtils.getExtension(file.getName()).toLowerCase();
                    long fileSizeMB = file.length() / (1024 * 1024);
                    
                    if (extension.equals("xlsb")) {
                        log.info("Found XLSB file: {} ({} MB). Note: XLSB format has limited support.", 
                                file.getName(), fileSizeMB);
                    } else {
                        log.debug("Found Excel file: {} ({} MB)", file.getName(), fileSizeMB);
                    }
                }
            }
        } else {
            log.warn("Excel directory '{}' does not exist", config.getExcelFolderPath());
            // Create the directory
            boolean created = excelDir.mkdirs();
            if (created) {
                log.info("Created Excel directory: {}", config.getExcelFolderPath());
            }
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
                break;
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
        Instant startTime = Instant.now();
        MappedByteBuffer mappedBuffer = null;
        
        try {
            // Set zip parameters to bypass zip bomb detection for this specific operation
            ZipSecureFile.setMinInflateRatio(0.0001); // More permissive ratio
            ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024); // 100MB max entry size
            
            String fileName = file.getName().toLowerCase();
            long fileSize = file.length() / (1024 * 1024); // Size in MB
            
            log.debug("Creating workbook for file: {} ({}MB)", file.getName(), fileSize);
            
            // Use memory-mapped files for better performance with large files
            if (config.isUseMemoryMapped() && fileSize > 5) { // Only for files > 5MB
                try {
                    log.debug("Using memory-mapped approach for large file: {}", file.getName());
                    fis.close(); // Close the input stream as we'll use memory mapping
                    
                    if (fileName.endsWith(".xlsx")) {
                        // For XLSX files, use OPC package with memory-mapped buffer
                        mappedBuffer = memoryMapper.createMemoryMappedBuffer(file);
                        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
                            return new XSSFWorkbook(pkg);
                        }
                    } else if (fileName.endsWith(".xls")) {
                        // For XLS files, use direct buffering
                        ByteBuffer buffer = memoryMapper.readFileToDirectBuffer(file, config.getBufferSize());
                        return new HSSFWorkbook(fis);
                    }
                    // For other formats, fall back to standard approach
                } catch (Exception e) {
                    log.warn("Memory-mapped approach failed, falling back to standard: {}", e.getMessage());
                    // Reopen the stream
                    fis = new FileInputStream(file);
                }
            }
            
            // Process by file type
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
            } else if (fileName.endsWith(".csv")) {
                // For CSV files, we create a wrapper that makes them appear as Excel workbooks
                log.info("CSV file detected: {}. Creating wrapper workbook.", file.getName());
                // This would require CSV workbook implementation - for now just throw exception
                throw new IOException("CSV files not yet supported directly. Please convert to Excel format.");
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
            metricsService.recordProcessingError("WorkbookCreation");
            throw new IOException("Failed to create workbook: " + e.getMessage(), e);
        } finally {
            Instant endTime = Instant.now();
            long elapsedTime = Duration.between(startTime, endTime).toMillis();
            log.debug("Workbook creation took {} ms for file: {}", elapsedTime, file.getName());
            
            // Release memory-mapped buffer if used
            if (mappedBuffer != null) {
                memoryMapper.releaseBuffer(mappedBuffer);
            }
        }
    }

    private List<ExcelData> extractDataFromSheet(Sheet sheet, String fileName) {
        Instant startTime = Instant.now();
        List<ExcelData> sheetData = new ArrayList<>();
        
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return sheetData;
        }
        
        try {
            // Get header row and determine all columns dynamically
            int firstRowNum = sheet.getFirstRowNum();
            Row headerRow = sheet.getRow(firstRowNum);
            Map<Integer, String> columnIndexToHeaderMap = new HashMap<>();
            Set<Integer> cellIndexes = new HashSet<>();
            int maxColumns = 0;
            
            // Phase 1: Analyze the sheet structure - determine used columns
            log.debug("Analyzing structure of sheet '{}' in file '{}'", sheet.getSheetName(), fileName);
            
            // First pass: determine the maximum number of columns across all rows
            // Use a sample of rows for better performance on very large sheets
            int lastRowNum = sheet.getLastRowNum();
            int rowCount = lastRowNum - firstRowNum + 1;
            
            // For very large sheets, sample rows instead of scanning all
            int sampleSize = Math.min(rowCount, 1000); // Sample at most 1000 rows
            int step = rowCount / sampleSize;
            if (step < 1) step = 1;
            
            // First analyze data rows to find all used columns
            for (int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex += step) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;
                
                short lastCellNum = row.getLastCellNum();
                if (lastCellNum > maxColumns) {
                    maxColumns = lastCellNum;
                }
                
                // Collect all cell indexes that have data
                for (Cell cell : row) {
                    int cellIndex = cell.getColumnIndex();
                    Object value = getCellValue(cell);
                    if (value != null && !value.toString().trim().isEmpty()) {
                        cellIndexes.add(cellIndex);
                    }
                }
            }
            
            log.debug("Sheet '{}' has maximum of {} columns with data in {} positions", 
                    sheet.getSheetName(), maxColumns, cellIndexes.size());
            
            // Phase 2: Extract headers
            List<String> headers = new ArrayList<>();
            
            // Deduplicate header names to ensure uniqueness
            Set<String> usedHeaderNames = new HashSet<>();
            
            if (headerRow != null) {
                // Get all headers from the first row
                for (int cellIndex = 0; cellIndex < maxColumns; cellIndex++) {
                    String headerValue;
                    Cell cell = headerRow.getCell(cellIndex);
                    
                    // If this column was used in any row, ensure it has a header
                    if (cellIndexes.contains(cellIndex) || cell != null) {
                        headerValue = getCellValueAsString(cell);
                        
                        // Handle empty or duplicate headers
                        if (headerValue == null || headerValue.trim().isEmpty()) {
                            headerValue = "Column_" + (cellIndex + 1);
                        }
                        
                        // Ensure header name uniqueness
                        String originalHeader = headerValue;
                        int suffix = 1;
                        while (usedHeaderNames.contains(headerValue.toLowerCase())) {
                            headerValue = originalHeader + "_" + suffix;
                            suffix++;
                        }
                        
                        usedHeaderNames.add(headerValue.toLowerCase());
                        columnIndexToHeaderMap.put(cellIndex, headerValue);
                    }
                }
            } else {
                // No header row found, generate default headers for all used columns
                for (Integer cellIndex : cellIndexes) {
                    String headerValue = "Column_" + (cellIndex + 1);
                    columnIndexToHeaderMap.put(cellIndex, headerValue);
                }
            }
            
            // Sort headers by column index
            headers = columnIndexToHeaderMap.entrySet().stream()
                .sorted(Map.Entry.comparingByKey())
                .map(Map.Entry::getValue)
                .collect(Collectors.toList());
            
            log.debug("Extracted {} column headers from sheet '{}': {}", 
                    headers.size(), sheet.getSheetName(), 
                    headers.size() > 10 ? headers.subList(0, 10) + "..." : headers);
            
            // Phase 3: Process data rows in batches for better memory efficiency
            int startRow = firstRowNum + 1; // Skip header
            if (headerRow == null) startRow = firstRowNum;
            
            List<Integer> rowIndexes = new ArrayList<>();
            for (int i = startRow; i <= lastRowNum; i++) {
                rowIndexes.add(i);
            }
            
            // Process in batches to reduce memory pressure
            int batchSize = config.getBatchSize();
            int totalRows = rowIndexes.size();
            int batches = (totalRows + batchSize - 1) / batchSize;
            
            log.debug("Processing {} rows in {} batches of size {}", 
                    totalRows, batches, batchSize);
            
            for (int batchIndex = 0; batchIndex < batches; batchIndex++) {
                int fromIndex = batchIndex * batchSize;
                int toIndex = Math.min(fromIndex + batchSize, totalRows);
                
                List<Integer> batchRowIndexes = rowIndexes.subList(fromIndex, toIndex);
                List<ExcelData> batchData = processBatch(sheet, fileName, batchRowIndexes, columnIndexToHeaderMap);
                sheetData.addAll(batchData);
            }
            
        } catch (Exception e) {
            log.error("Error extracting data from sheet '{}' in file '{}'", sheet.getSheetName(), fileName, e);
            metricsService.recordProcessingError("SheetProcessing");
        }
        
        Instant endTime = Instant.now();
        long elapsedTime = Duration.between(startTime, endTime).toMillis();
        log.info("Extracted {} rows from sheet '{}' in {} ms", 
                sheetData.size(), sheet.getSheetName(), elapsedTime);
        
        // Record metrics
        metricsService.recordSheetProcessingTime(sheet.getSheetName(), elapsedTime);
        metricsService.recordRowsProcessed(sheetData.size());
        
        return sheetData;
    }
    
    /**
     * Process a batch of rows from an Excel sheet.
     * 
     * @param sheet The sheet to process
     * @param fileName The name of the file
     * @param rowIndexes The indexes of the rows to process
     * @param columnIndexToHeaderMap Mapping of column indexes to header names
     * @return List of ExcelData objects for the batch
     */
    private List<ExcelData> processBatch(Sheet sheet, String fileName, List<Integer> rowIndexes, 
                                        Map<Integer, String> columnIndexToHeaderMap) {
        List<ExcelData> batchData = new ArrayList<>();
        
        for (Integer rowIndex : rowIndexes) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            
            Map<String, Object> rowData = new HashMap<>();
            boolean hasData = false;
            
            // Extract data for all mapped columns
            for (Map.Entry<Integer, String> entry : columnIndexToHeaderMap.entrySet()) {
                int cellIndex = entry.getKey();
                String header = entry.getValue();
                
                Cell cell = row.getCell(cellIndex);
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
                excelData.setRowNumber(rowIndex + 1); // 1-based row numbers for user display
                excelData.setData(rowData);
                excelData.setExtractedAt(LocalDateTime.now().format(FORMATTER));
                
                batchData.add(excelData);
            }
        }
        
        return batchData;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                    
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // Return dates in a standard format
                        try {
                            return cell.getDateCellValue().toString();
                        } catch (Exception e) {
                            // Fall back to numeric value if date conversion fails
                            return cell.getNumericCellValue();
                        }
                    } else {
                        double value = cell.getNumericCellValue();
                        // Check if it's an integer value stored as double
                        if (value == Math.floor(value) && !Double.isInfinite(value)) {
                            // It's an integer
                            if (value >= Long.MIN_VALUE && value <= Long.MAX_VALUE) {
                                return (long) value;
                            }
                        }
                        return value;
                    }
                    
                case BOOLEAN:
                    return cell.getBooleanCellValue();
                    
                case FORMULA:
                    // Try to evaluate the formula
                    try {
                        switch (cell.getCachedFormulaResultType()) {
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    return cell.getDateCellValue().toString();
                                } else {
                                    double value = cell.getNumericCellValue();
                                    if (value == Math.floor(value) && !Double.isInfinite(value)) {
                                        if (value >= Long.MIN_VALUE && value <= Long.MAX_VALUE) {
                                            return (long) value;
                                        }
                                    }
                                    return value;
                                }
                            case STRING:
                                return cell.getStringCellValue();
                            case BOOLEAN:
                                return cell.getBooleanCellValue();
                            case ERROR:
                                return "#ERROR";
                            default:
                                return cell.getCellFormula();
                        }
                    } catch (Exception e) {
                        // If formula evaluation fails, return the formula itself
                        try {
                            return "="+cell.getCellFormula();
                        } catch (Exception ex) {
                            return "#FORMULA_ERROR";
                        }
                    }
                    
                case BLANK:
                    return null;
                    
                case ERROR:
                    return "#ERROR:" + cell.getErrorCellValue();
                    
                default:
                    return null;
            }
        } catch (Exception e) {
            // Safe fallback for any cell reading errors
            log.debug("Error reading cell value: {}", e.getMessage());
            return "#ERROR_READING_CELL";
        }
    }

    private String getCellValueAsString(Cell cell) {
        Object value = getCellValue(cell);
        if (value == null) {
            return "";
        }
        
        // Format numeric values to avoid scientific notation
        if (value instanceof Double) {
            double doubleValue = (Double) value;
            // Check if it's a whole number
            if (doubleValue == Math.floor(doubleValue) && !Double.isInfinite(doubleValue)) {
                return String.format("%.0f", doubleValue);
            } else {
                return String.valueOf(value);
            }
        }
        
        return String.valueOf(value);
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
    
    @Override
    public List<ExcelData> transformDataByCombiningColumns(List<ExcelData> data, List<String> sourceColumns, 
                                                         String targetColumn, String separator) {
        log.info("Transforming data by combining columns: {} -> {}", sourceColumns, targetColumn);
        List<ExcelData> transformedData = new ArrayList<>();
        
        if (separator == null) {
            separator = "";
        }
        
        for (ExcelData excelData : data) {
            // Create a deep copy of the original data
            ExcelData newData = new ExcelData();
            newData.setFileName(excelData.getFileName());
            newData.setSheetName(excelData.getSheetName());
            newData.setRowNumber(excelData.getRowNumber());
            newData.setExtractedAt(excelData.getExtractedAt());
            
            Map<String, Object> rowData = new HashMap<>(excelData.getData());
            
            // Build combined value from source columns
            StringBuilder combinedValue = new StringBuilder();
            boolean firstColumn = true;
            
            for (String columnName : sourceColumns) {
                Object columnValue = rowData.get(columnName);
                if (columnValue != null) {
                    if (!firstColumn && !separator.isEmpty()) {
                        combinedValue.append(separator);
                    }
                    combinedValue.append(columnValue.toString().trim());
                    firstColumn = false;
                }
            }
            
            // Add the new combined column
            rowData.put(targetColumn, combinedValue.toString());
            newData.setData(rowData);
            
            transformedData.add(newData);
        }
        
        log.info("Transformation complete. Created new column '{}' in {} records", 
                targetColumn, transformedData.size());
        
        return transformedData;
    }
    
    @Override
    public String createTransformedExcelFile(List<ExcelData> data, 
                                          Map<String, List<String>> transformationMap, 
                                          Map<String, String> separatorMap, 
                                          String outputFileName, 
                                          boolean includeOriginalColumns) {
        log.info("Creating transformed Excel file with {} column transformations", transformationMap.size());
        
        if (separatorMap == null) {
            separatorMap = new HashMap<>();
        }
        
        // Create a copy of the original data with transformed columns
        List<Map<String, Object>> transformedRows = new ArrayList<>();
        
        for (ExcelData excelData : data) {
            // Start with empty row or original data based on includeOriginalColumns flag
            Map<String, Object> transformedRow = includeOriginalColumns ? 
                new HashMap<>(excelData.getData()) : new HashMap<>();
            
            // Apply each transformation
            for (Map.Entry<String, List<String>> transformation : transformationMap.entrySet()) {
                String targetColumn = transformation.getKey();
                List<String> sourceColumns = transformation.getValue();
                String separator = separatorMap.getOrDefault(targetColumn, "");
                
                // Build combined value from source columns
                StringBuilder combinedValue = new StringBuilder();
                boolean firstColumn = true;
                
                for (String columnName : sourceColumns) {
                    Object columnValue = excelData.getData().get(columnName);
                    if (columnValue != null) {
                        if (!firstColumn && !separator.isEmpty()) {
                            combinedValue.append(separator);
                        }
                        combinedValue.append(columnValue.toString().trim());
                        firstColumn = false;
                    }
                }
                
                // Add the transformed column to the row
                transformedRow.put(targetColumn, combinedValue.toString());
            }
            
            transformedRows.add(transformedRow);
        }
        
        // Ensure output file has proper extension
        if (!outputFileName.toLowerCase().endsWith(".xlsx") && 
            !outputFileName.toLowerCase().endsWith(".xls")) {
            outputFileName += ".xlsx";
        }
        
        // Create directory if it doesn't exist
        File outputDir = new File(TEMP_FOLDER_PATH);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }
        
        // Create full path for output file
        String outputFilePath = TEMP_FOLDER_PATH + File.separator + outputFileName;
        File outputFile = new File(outputFilePath);
        
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Transformed Data");
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            
            // Get all column names from the first row
            Set<String> columnNames = new HashSet<>();
            if (!transformedRows.isEmpty()) {
                columnNames.addAll(transformedRows.get(0).keySet());
            }
            
            // Create header cells
            int cellIndex = 0;
            for (String columnName : columnNames) {
                Cell cell = headerRow.createCell(cellIndex++);
                cell.setCellValue(columnName);
            }
            
            // Create data rows
            for (int rowIndex = 0; rowIndex < transformedRows.size(); rowIndex++) {
                Row row = sheet.createRow(rowIndex + 1);  // +1 to account for header
                Map<String, Object> rowData = transformedRows.get(rowIndex);
                
                cellIndex = 0;
                for (String columnName : columnNames) {
                    Cell cell = row.createCell(cellIndex++);
                    Object value = rowData.get(columnName);
                    
                    if (value != null) {
                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else if (value instanceof Date) {
                            cell.setCellValue((Date) value);
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
            }
            
            // Auto-size columns for better readability
            for (int i = 0; i < columnNames.size(); i++) {
                sheet.autoSizeColumn(i);
            }
            
            // Write to file
            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                workbook.write(fileOut);
            }
            
            log.info("Successfully created transformed Excel file at {}", outputFile.getAbsolutePath());
            return outputFile.getAbsolutePath();
            
        } catch (Exception e) {
            log.error("Error creating transformed Excel file: {}", e.getMessage(), e);
            throw new RuntimeException("Failed to create transformed Excel file", e);
        }
    }
}
