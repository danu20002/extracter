package com.jnj.extracter.api.controller;

import com.jnj.extracter.service.contract.TransformService;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * REST controller for Excel transformation operations.
 */
@RestController
@RequestMapping("/api/transform")
@RequiredArgsConstructor
@Slf4j
@Tag(name = "Transform API", description = "API for Excel data transformation operations")
public class TransformController {

    private final TransformService transformService;

    /**
     * Transform Excel data by combining columns.
     * 
     * @param sourceColumns List of source column names
     * @param targetColumn Name of the target column
     * @param separator Separator for combining values (optional)
     * @param fileName Name of the file to generate
     * @param includeOriginalColumns Whether to include original columns
     * @return Path to the generated file
     */
    @PostMapping("/combine")
    @Operation(summary = "Transform Excel data by combining columns")
    public ResponseEntity<Map<String, Object>> transformByCombining(
            @RequestParam List<String> sourceColumns,
            @RequestParam String targetColumn,
            @RequestParam(required = false, defaultValue = "") String separator,
            @RequestParam String fileName,
            @RequestParam(required = false, defaultValue = "true") boolean includeOriginalColumns) {
        
        try {
            log.info("Transforming data by combining columns: {} -> {}", sourceColumns, targetColumn);
            
            // Create transformation map
            Map<String, List<String>> transformationMap = new HashMap<>();
            transformationMap.put(targetColumn, sourceColumns);
            
            // Create separator map
            Map<String, String> separatorMap = new HashMap<>();
            separatorMap.put(targetColumn, separator);
            
            // Create a sample row for demonstration
            Map<String, Object> sampleData = new HashMap<>();
            for (String col : sourceColumns) {
                sampleData.put(col, "Sample " + col);
            }
            
            // Get example data from the first Excel file
            // Note: In a real implementation, we would get data from a specific file
            List<Map<String, Object>> testData = List.of(sampleData);
            
            // Generate the file
            String outputPath = transformService.createTransformedExcelFile(
                    testData, 
                    transformationMap, 
                    separatorMap,
                    fileName,
                    includeOriginalColumns);
            
            // Return result
            Map<String, Object> result = new HashMap<>();
            result.put("status", "success");
            result.put("message", "Transformation completed successfully");
            result.put("outputPath", outputPath);
            
            return ResponseEntity.ok(result);
            
        } catch (Exception e) {
            log.error("Error during transformation", e);
            
            Map<String, Object> error = new HashMap<>();
            error.put("status", "error");
            error.put("message", "Transformation failed: " + e.getMessage());
            
            return ResponseEntity.internalServerError().body(error);
        }
    }

    /**
     * Download a transformed Excel file.
     * 
     * @param fileName Name of the file to download
     * @return The file as a response
     */
    @GetMapping("/download/{fileName}")
    @Operation(summary = "Download a transformed Excel file")
    public ResponseEntity<ByteArrayResource> downloadFile(@PathVariable String fileName) {
        try {
            String filePath = "excel/temp/" + fileName;
            if (!fileName.toLowerCase().endsWith(".xlsx")) {
                filePath += ".xlsx";
            }
            
            Path path = Paths.get(filePath);
            byte[] data = Files.readAllBytes(path);
            
            ByteArrayResource resource = new ByteArrayResource(data);
            
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .contentLength(data.length)
                    .body(resource);
                    
        } catch (IOException e) {
            log.error("Error downloading file: {}", fileName, e);
            return ResponseEntity.notFound().build();
        }
    }
}
