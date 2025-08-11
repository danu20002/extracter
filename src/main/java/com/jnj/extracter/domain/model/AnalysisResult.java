package com.jnj.extracter.domain.model;

import lombok.Data;

import java.util.Map;

/**
 * Represents the result of analyzing an Excel file.
 * Contains various statistics about the file, its sheets, and columns.
 */
@Data
public class AnalysisResult {
    private String filename;
    private int sheetCount;
    
    // File-level statistics
    private Map<String, Object> fileStatistics;
    
    // Sheet-level statistics, keyed by sheet name
    private Map<String, Map<String, Object>> sheetStatistics;
    
    // Column-level statistics, keyed by sheet name and column name
    private Map<String, Map<String, Map<String, Object>>> columnStatistics;
}
