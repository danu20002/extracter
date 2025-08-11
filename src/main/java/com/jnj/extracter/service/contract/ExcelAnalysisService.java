package com.jnj.extracter.service.contract;

import com.jnj.extracter.domain.model.ExcelData;

import java.util.List;
import java.util.Map;

/**
 * Service interface for analyzing Excel data.
 * This interface separates data analysis functionality from 
 * the Excel reading and processing logic.
 */
public interface ExcelAnalysisService {
    
    /**
     * Calculate summary statistics for Excel data.
     * 
     * @param data The Excel data to analyze
     * @return Map of summary statistics
     */
    Map<String, Object> calculateSummaryStatistics(List<ExcelData> data);
    
    /**
     * Analyze the distribution of values in a specific column.
     * 
     * @param data The Excel data
     * @param columnName The column to analyze
     * @return Map containing distribution information
     */
    Map<String, Object> analyzeColumnDistribution(List<ExcelData> data, String columnName);
    
    /**
     * Calculate correlations between numeric columns.
     * 
     * @param data The Excel data
     * @return Matrix of correlation coefficients
     */
    Map<String, Map<String, Double>> calculateCorrelations(List<ExcelData> data);
    
    /**
     * Detect outliers in numeric columns.
     * 
     * @param data The Excel data
     * @return Map of column names to lists of outlier values and their row indices
     */
    Map<String, List<Map<String, Object>>> detectOutliers(List<ExcelData> data);
    
    /**
     * Validate data against business rules.
     * 
     * @param data The Excel data
     * @param validationRules Map of column names to validation rules
     * @return Map of column names to validation results
     */
    Map<String, List<Map<String, Object>>> validateData(
            List<ExcelData> data, 
            Map<String, String> validationRules);
}
