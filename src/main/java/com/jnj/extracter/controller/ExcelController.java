package com.jnj.extracter.controller;

import com.jnj.extracter.entity.ExcelData;
import com.jnj.extracter.entity.ExcelProcessingResult;
import com.jnj.extracter.service.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelController {

    private final ExcelService excelService;

    /**
     * Extract data from all Excel files in the excel folder
     */
    @GetMapping("/extract-all")
    public ResponseEntity<List<ExcelProcessingResult>> extractAllExcelFiles() {
        List<ExcelProcessingResult> results = excelService.extractAllExcelFiles();
        return ResponseEntity.ok(results);
    }

    /**
     * Get list of available Excel files
     */
    @GetMapping("/files")
    public ResponseEntity<List<String>> getExcelFiles() {
        List<File> files = excelService.getExcelFiles();
        List<String> fileNames = files.stream()
                .map(File::getName)
                .toList();
        return ResponseEntity.ok(fileNames);
    }

    /**
     * Extract data from a specific Excel file
     */
    @GetMapping("/extract/{fileName}")
    public ResponseEntity<ExcelProcessingResult> extractSpecificFile(@PathVariable String fileName) {
        List<File> files = excelService.getExcelFiles();
        File targetFile = files.stream()
                .filter(file -> file.getName().equals(fileName))
                .findFirst()
                .orElse(null);

        if (targetFile == null) {
            ExcelProcessingResult errorResult = new ExcelProcessingResult();
            errorResult.setFileName(fileName);
            errorResult.setSuccess(false);
            errorResult.setMessage("File not found: " + fileName);
            return ResponseEntity.notFound().build();
        }

        ExcelProcessingResult result = excelService.extractExcelFile(targetFile);
        return ResponseEntity.ok(result);
    }

    /**
     * Extract data from a specific sheet in a file
     */
    @GetMapping("/extract/{fileName}/{sheetName}")
    public ResponseEntity<List<ExcelData>> extractSheetData(
            @PathVariable String fileName,
            @PathVariable String sheetName) {
        
        List<File> files = excelService.getExcelFiles();
        File targetFile = files.stream()
                .filter(file -> file.getName().equals(fileName))
                .findFirst()
                .orElse(null);

        if (targetFile == null) {
            return ResponseEntity.notFound().build();
        }

        List<ExcelData> data = excelService.extractSheetData(targetFile, sheetName);
        return ResponseEntity.ok(data);
    }

    /**
     * Perform operations on extracted data
     */
    @PostMapping("/operations/{operation}")
    public ResponseEntity<Map<String, Object>> performDataOperations(
            @PathVariable String operation,
            @RequestBody List<ExcelData> data) {
        
        Map<String, Object> result = excelService.performDataOperations(data, operation);
        return ResponseEntity.ok(result);
    }

    /**
     * Get data summary for all extracted data
     */
    @GetMapping("/summary")
    public ResponseEntity<Map<String, Object>> getDataSummary() {
        List<ExcelProcessingResult> results = excelService.extractAllExcelFiles();
        List<ExcelData> allData = results.stream()
                .flatMap(result -> result.getExtractedData().stream())
                .toList();
        
        Map<String, Object> summary = excelService.getDataSummary(allData);
        return ResponseEntity.ok(summary);
    }

    /**
     * Perform specific operation on all extracted data
     */
    @GetMapping("/operations/{operation}")
    public ResponseEntity<Map<String, Object>> performOperationOnAllData(@PathVariable String operation) {
        List<ExcelProcessingResult> results = excelService.extractAllExcelFiles();
        List<ExcelData> allData = results.stream()
                .flatMap(result -> result.getExtractedData().stream())
                .toList();
        
        Map<String, Object> result = excelService.performDataOperations(allData, operation);
        return ResponseEntity.ok(result);
    }

    /**
     * Transform data by combining columns
     */
    @PostMapping("/transform/combine-columns")
    public ResponseEntity<Map<String, Object>> transformByCombiningColumns(
            @RequestBody Map<String, Object> requestBody) {
        
        try {
            @SuppressWarnings("unchecked")
            List<ExcelData> data = (List<ExcelData>) requestBody.get("data");
            
            @SuppressWarnings("unchecked")
            List<String> sourceColumns = (List<String>) requestBody.get("sourceColumns");
            
            String targetColumn = (String) requestBody.get("targetColumn");
            String separator = (String) requestBody.get("separator");
            
            if (data == null || sourceColumns == null || targetColumn == null || sourceColumns.isEmpty()) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "Invalid request. Required fields: data, sourceColumns, targetColumn");
                return ResponseEntity.badRequest().body(error);
            }
            
            List<ExcelData> transformedData = excelService.transformDataByCombiningColumns(
                    data, sourceColumns, targetColumn, separator);
            
            Map<String, Object> response = new HashMap<>();
            response.put("success", true);
            response.put("transformedData", transformedData);
            response.put("recordCount", transformedData.size());
            response.put("sourceColumns", sourceColumns);
            response.put("targetColumn", targetColumn);
            
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            Map<String, Object> error = new HashMap<>();
            error.put("error", "Error during transformation: " + e.getMessage());
            return ResponseEntity.badRequest().body(error);
        }
    }
    
    /**
     * Create a new Excel file with multiple transformed columns
     */
    @PostMapping("/transform/create-excel")
    public ResponseEntity<Map<String, Object>> createTransformedExcelFile(
            @RequestBody Map<String, Object> requestBody) {
        
        try {
            @SuppressWarnings("unchecked")
            List<ExcelData> data = (List<ExcelData>) requestBody.get("data");
            
            @SuppressWarnings("unchecked")
            Map<String, List<String>> transformationMap = (Map<String, List<String>>) requestBody.get("transformations");
            
            @SuppressWarnings("unchecked")
            Map<String, String> separatorMap = (Map<String, String>) requestBody.get("separators");
            
            String outputFileName = (String) requestBody.get("outputFileName");
            Boolean includeOriginalColumns = (Boolean) requestBody.get("includeOriginalColumns");
            
            if (data == null || transformationMap == null || transformationMap.isEmpty() || outputFileName == null) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "Invalid request. Required fields: data, transformations, outputFileName");
                return ResponseEntity.badRequest().body(error);
            }
            
            if (includeOriginalColumns == null) {
                includeOriginalColumns = false;
            }
            
            String filePath = excelService.createTransformedExcelFile(
                    data, transformationMap, separatorMap, outputFileName, includeOriginalColumns);
            
            Map<String, Object> response = new HashMap<>();
            response.put("success", true);
            response.put("filePath", filePath);
            response.put("fileName", new File(filePath).getName());
            response.put("transformationCount", transformationMap.size());
            response.put("recordCount", data.size());
            
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            Map<String, Object> error = new HashMap<>();
            error.put("error", "Error creating transformed Excel file: " + e.getMessage());
            return ResponseEntity.badRequest().body(error);
        }
    }
    
    /**
     * Transform data for a specific file by combining columns
     */
    @PostMapping("/transform/file/{fileName}")
    public ResponseEntity<Map<String, Object>> transformFileDataByCombiningColumns(
            @PathVariable String fileName,
            @RequestBody Map<String, Object> requestBody) {
        
        try {
            List<File> files = excelService.getExcelFiles();
            File targetFile = files.stream()
                    .filter(file -> file.getName().equals(fileName))
                    .findFirst()
                    .orElse(null);
    
            if (targetFile == null) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "File not found: " + fileName);
                return ResponseEntity.notFound().build();
            }
            
            ExcelProcessingResult result = excelService.extractExcelFile(targetFile);
            
            if (!result.isSuccess()) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "Failed to extract data from file: " + result.getMessage());
                return ResponseEntity.badRequest().body(error);
            }
            
            @SuppressWarnings("unchecked")
            List<String> sourceColumns = (List<String>) requestBody.get("sourceColumns");
            
            String targetColumn = (String) requestBody.get("targetColumn");
            String separator = (String) requestBody.get("separator");
            
            if (sourceColumns == null || targetColumn == null || sourceColumns.isEmpty()) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "Invalid request. Required fields: sourceColumns, targetColumn");
                return ResponseEntity.badRequest().body(error);
            }
            
            List<ExcelData> transformedData = excelService.transformDataByCombiningColumns(
                    result.getExtractedData(), sourceColumns, targetColumn, separator);
            
            Map<String, Object> response = new HashMap<>();
            response.put("success", true);
            response.put("transformedData", transformedData);
            response.put("recordCount", transformedData.size());
            response.put("fileName", fileName);
            response.put("sourceColumns", sourceColumns);
            response.put("targetColumn", targetColumn);
            
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            Map<String, Object> error = new HashMap<>();
            error.put("error", "Error during transformation: " + e.getMessage());
            return ResponseEntity.badRequest().body(error);
        }
    }
    
    /**
     * Create a new Excel file with multiple column transformations from a source file
     */
    @PostMapping("/transform/file/{fileName}/create-excel")
    public ResponseEntity<Map<String, Object>> createTransformedExcelFromFile(
            @PathVariable String fileName,
            @RequestBody Map<String, Object> requestBody) {
        
        try {
            List<File> files = excelService.getExcelFiles();
            File targetFile = files.stream()
                    .filter(file -> file.getName().equals(fileName))
                    .findFirst()
                    .orElse(null);
    
            if (targetFile == null) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "File not found: " + fileName);
                return ResponseEntity.notFound().build();
            }
            
            // Extract data from file
            ExcelProcessingResult result = excelService.extractExcelFile(targetFile);
            
            if (!result.isSuccess()) {
                Map<String, Object> error = new HashMap<>();
                error.put("error", "Failed to extract data from file: " + result.getMessage());
                return ResponseEntity.badRequest().body(error);
            }
            
            @SuppressWarnings("unchecked")
            Map<String, List<String>> transformationMap = (Map<String, List<String>>) requestBody.get("transformations");
            
            @SuppressWarnings("unchecked")
            Map<String, String> separatorMap = (Map<String, String>) requestBody.get("separators");
            
            String outputFileName = (String) requestBody.get("outputFileName");
            Boolean includeOriginalColumns = (Boolean) requestBody.get("includeOriginalColumns");
            
            // Default output filename if not provided
            if (outputFileName == null || outputFileName.trim().isEmpty()) {
                String baseName = fileName.substring(0, fileName.lastIndexOf('.'));
                outputFileName = baseName + "_transformed.xlsx";
            }
            
            if (includeOriginalColumns == null) {
                includeOriginalColumns = false;
            }
            
            // Create transformed Excel file
            String filePath = excelService.createTransformedExcelFile(
                    result.getExtractedData(), 
                    transformationMap, 
                    separatorMap, 
                    outputFileName, 
                    includeOriginalColumns);
            
            Map<String, Object> response = new HashMap<>();
            response.put("success", true);
            response.put("filePath", filePath);
            response.put("fileName", new File(filePath).getName());
            response.put("sourceFile", fileName);
            response.put("transformationCount", transformationMap.size());
            response.put("recordCount", result.getExtractedData().size());
            
            return ResponseEntity.ok(response);
        } catch (Exception e) {
            Map<String, Object> error = new HashMap<>();
            error.put("error", "Error creating transformed Excel file: " + e.getMessage());
            return ResponseEntity.badRequest().body(error);
        }
    }

    /**
     * Health check endpoint
     */
    @GetMapping("/health")
    public ResponseEntity<Map<String, Object>> healthCheck() {
        List<File> files = excelService.getExcelFiles();
        
        // Check for XLSB files and provide warnings
        List<String> warnings = new ArrayList<>();
        int xlsbCount = 0;
        
        for (File file : files) {
            if (file.getName().toLowerCase().endsWith(".xlsb")) {
                xlsbCount++;
            }
        }
        
        if (xlsbCount > 0) {
            warnings.add(xlsbCount + " XLSB file(s) detected. XLSB format has limited support.");
        }
        
        Map<String, Object> health = new HashMap<>();
        health.put("status", "UP");
        health.put("excelFolderExists", new File("excel").exists());
        health.put("availableExcelFiles", files.size());
        health.put("fileNames", files.stream().map(File::getName).toList());
        
        if (!warnings.isEmpty()) {
            health.put("warnings", warnings);
        }
        
        return ResponseEntity.ok(health);
    }
}
