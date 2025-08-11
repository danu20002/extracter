package com.jnj.extracter.api.controller;

import com.jnj.extracter.api.exception.ResourceNotFoundException;
import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.service.contract.ExcelService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * REST API controller for Excel operations.
 */
@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
@Slf4j
@Tag(name = "Excel API", description = "API for Excel file operations")
public class ExcelApiController {

    private final ExcelService excelService;
    
    /**
     * Get sheet names for a file.
     * 
     * @param fileName Name of the Excel file
     * @return Map containing sheet names and file metadata
     */
    @GetMapping("/sheets/{fileName}")
    @Operation(summary = "Get sheet names for a file", description = "Returns all sheet names in the specified Excel file")
    public ResponseEntity<Map<String, Object>> getSheetNames(@PathVariable String fileName) {
        try {
            log.info("Getting sheet names for file: {}", fileName);
            
            // Find the file
            List<File> files = excelService.getExcelFiles();
            File targetFile = files.stream()
                .filter(f -> f.getName().equalsIgnoreCase(fileName))
                .findFirst()
                .orElseThrow(() -> new ResourceNotFoundException("Excel file", fileName));
            
            // Get sheet names
            List<String> sheetNames = excelService.getSheetNames(targetFile);
            
            // Return result
            Map<String, Object> result = Map.of(
                "fileName", targetFile.getName(),
                "filePath", targetFile.getPath(),
                "fileSize", targetFile.length(),
                "lastModified", new java.util.Date(targetFile.lastModified()).toString(),
                "sheetNames", sheetNames
            );
            
            return ResponseEntity.ok(result);
        } catch (ResourceNotFoundException e) {
            log.warn("File not found: {}", fileName);
            throw e;
        } catch (Exception e) {
            log.error("Error getting sheet names for file: {}", fileName, e);
            return ResponseEntity.internalServerError().body(
                Collections.singletonMap("error", "Failed to get sheet names: " + e.getMessage())
            );
        }
    }
}
