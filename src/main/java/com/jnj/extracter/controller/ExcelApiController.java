package com.jnj.extracter.controller;

import com.jnj.extracter.service.ExcelService;
import lombok.RequiredArgsConstructor;
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
 * REST API for Excel operations
 */
@RestController
@RequestMapping("/api/excel")
@RequiredArgsConstructor
public class ExcelApiController {

    private final ExcelService excelService;
    
    /**
     * Get sheet names for a file
     */
    @GetMapping("/sheets/{fileName}")
    public ResponseEntity<Map<String, Object>> getSheetNames(@PathVariable String fileName) {
        try {
            List<File> files = excelService.getExcelFiles();
            File targetFile = files.stream()
                    .filter(file -> file.getName().equals(fileName))
                    .findFirst()
                    .orElse(null);
                    
            if (targetFile == null) {
                return ResponseEntity.notFound().build();
            }
            
            List<String> sheetNames = excelService.getSheetNames(targetFile);
            return ResponseEntity.ok(Map.of(
                "fileName", fileName,
                "sheetNames", sheetNames
            ));
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of(
                "error", e.getMessage()
            ));
        }
    }
}
