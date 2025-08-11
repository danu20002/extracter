package com.jnj.extracter.service.impl;

import com.jnj.extracter.api.exception.ExcelProcessingException;
import com.jnj.extracter.api.exception.ResourceNotFoundException;
import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.ExcelFileInfo;
import com.jnj.extracter.domain.model.ExcelProcessingResult;
import com.jnj.extracter.domain.model.SheetData;
import com.jnj.extracter.service.contract.ExcelReaderService;
import com.jnj.extracter.service.contract.ExcelService;
import com.jnj.extracter.service.contract.MetricsService;
import com.jnj.extracter.service.contract.StorageService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

/**
 * Implementation of ExcelService that processes Excel files.
 */
@Service
@Slf4j
public class ExcelServiceImpl implements ExcelService {
    
    private final ExcelReaderService excelReaderService;
    private final StorageService storageService;
    private final MetricsService metricsService;
    
    // Cache to store processed excel data
    private final Map<String, ExcelData> excelDataCache = new ConcurrentHashMap<>();
    
    public ExcelServiceImpl(
            ExcelReaderService excelReaderService, 
            StorageService storageService,
            MetricsService metricsService) {
        this.excelReaderService = excelReaderService;
        this.storageService = storageService;
        this.metricsService = metricsService;
        log.info("ExcelServiceImpl initialized");
    }

    @Override
    public List<ExcelFileInfo> getAllExcelFiles() {
        log.debug("Retrieving all Excel files");
        try {
            List<Path> filePaths = storageService.loadAll();
            List<ExcelFileInfo> fileInfos = new ArrayList<>();
            
            for (Path filePath : filePaths) {
                String filename = filePath.getFileName().toString();
                if (isExcelFile(filename)) {
                    ExcelFileInfo fileInfo = new ExcelFileInfo();
                    fileInfo.setFilename(filename);
                    fileInfo.setPath(filePath.toString());
                    fileInfo.setSize(storageService.getFileSize(filePath));
                    fileInfo.setLastModified(storageService.getLastModifiedTime(filePath).toString());
                    fileInfos.add(fileInfo);
                }
            }
            
            log.info("Found {} Excel files", fileInfos.size());
            return fileInfos;
        } catch (IOException e) {
            log.error("Error retrieving Excel files", e);
            throw new ExcelProcessingException("Failed to retrieve Excel files", e);
        }
    }
    
    @Override
    public List<File> getExcelFiles() {
        log.debug("Retrieving Excel files as File objects");
        try {
            List<Path> filePaths = storageService.loadAll();
            List<File> files = filePaths.stream()
                    .map(path -> path.toFile())
                    .filter(file -> isExcelFile(file.getName()))
                    .collect(Collectors.toList());
            
            log.info("Found {} Excel files", files.size());
            return files;
        } catch (IOException e) {
            log.error("Error retrieving Excel files", e);
            throw new ExcelProcessingException("Failed to retrieve Excel files", e);
        }
    }

    @Override
    public List<String> getSheetNames(File file) throws IOException {
        log.debug("Getting sheet names for file: {}", file.getName());
        
        try (Workbook workbook = WorkbookFactory.create(file)) {
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
            return sheetNames;
        } catch (IOException e) {
            log.error("Error reading Excel file: {}", file.getName(), e);
            throw e;
        }
    }
    
    @Override
    public ExcelData getExcelData(String filename) {
        log.debug("Retrieving Excel data for file: {}", filename);
        
        // Check cache first
        if (excelDataCache.containsKey(filename)) {
            log.debug("Retrieved Excel data from cache for file: {}", filename);
            return excelDataCache.get(filename);
        }
        
        try {
            long startTime = System.currentTimeMillis();
            Path filePath = storageService.load(filename);
            if (filePath == null) {
                throw new ResourceNotFoundException("File not found: " + filename);
            }
            
            ExcelData excelData = excelReaderService.readExcelFile(filePath);
            
            // Record metrics
            long processingTime = System.currentTimeMillis() - startTime;
            metricsService.recordFileProcessingTime(filename, processingTime);
            metricsService.incrementFilesProcessed();
            
            // Cache the data
            excelDataCache.put(filename, excelData);
            
            log.info("Excel data processed for file: {} in {}ms", filename, processingTime);
            return excelData;
        } catch (IOException e) {
            log.error("Error reading Excel file: {}", filename, e);
            metricsService.incrementProcessingErrors("io_error");
            throw new ExcelProcessingException("Failed to read Excel file: " + filename, e);
        }
    }
    
    @Override
    public SheetData getSheetData(String filename, String sheetName) {
        log.debug("Retrieving sheet data for file: {}, sheet: {}", filename, sheetName);
        
        ExcelData excelData = getExcelData(filename);
        
        Optional<SheetData> sheetDataOpt = excelData.getSheets().stream()
                .filter(sheet -> sheet.getName().equals(sheetName))
                .findFirst();
        
        if (sheetDataOpt.isPresent()) {
            log.debug("Found sheet data for sheet: {}", sheetName);
            return sheetDataOpt.get();
        } else {
            log.warn("Sheet not found: {} in file: {}", sheetName, filename);
            throw new ResourceNotFoundException("Sheet not found: " + sheetName + " in file: " + filename);
        }
    }
    
    @Override
    public ExcelProcessingResult processExcelFile(MultipartFile file) {
        log.debug("Processing uploaded Excel file: {}", file.getOriginalFilename());
        
        try {
            long startTime = System.currentTimeMillis();
            
            // Store the file
            String filename = storageService.store(file);
            Path filePath = storageService.load(filename);
            
            // Process the file
            ExcelData excelData = excelReaderService.readExcelFile(filePath);
            
            // Cache the data
            excelDataCache.put(filename, excelData);
            
            // Prepare result
            ExcelProcessingResult result = new ExcelProcessingResult();
            result.setFilename(filename);
            result.setSheetCount(excelData.getSheets().size());
            result.setTotalRowCount(excelData.getSheets().stream()
                    .mapToInt(sheet -> sheet.getRows().size())
                    .sum());
            
            // Record metrics
            long processingTime = System.currentTimeMillis() - startTime;
            metricsService.recordFileProcessingTime(filename, processingTime);
            metricsService.incrementFilesProcessed();
            metricsService.incrementRowsProcessed(result.getTotalRowCount());
            
            log.info("Successfully processed file: {} with {} sheets and {} rows in {}ms", 
                    filename, result.getSheetCount(), result.getTotalRowCount(), processingTime);
                    
            return result;
        } catch (IOException e) {
            log.error("Error processing Excel file: {}", file.getOriginalFilename(), e);
            metricsService.incrementProcessingErrors("processing_error");
            throw new ExcelProcessingException("Failed to process Excel file: " + file.getOriginalFilename(), e);
        }
    }
    
    @Override
    public void clearCache() {
        log.info("Clearing Excel data cache with {} entries", excelDataCache.size());
        excelDataCache.clear();
    }
    
    @Override
    public void removeFromCache(String filename) {
        log.debug("Removing file from cache: {}", filename);
        excelDataCache.remove(filename);
    }
    
    private boolean isExcelFile(String filename) {
        return filename.toLowerCase().endsWith(".xlsx") || 
               filename.toLowerCase().endsWith(".xls");
    }
}
