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
