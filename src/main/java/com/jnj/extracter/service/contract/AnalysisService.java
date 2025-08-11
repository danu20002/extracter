package com.jnj.extracter.service.contract;

import com.jnj.extracter.domain.model.AnalysisResult;

/**
 * Service interface for analyzing Excel files and extracting statistical information.
 */
public interface AnalysisService {
    
    /**
     * Analyzes an Excel file and returns statistical information.
     * 
     * @param filename The name of the Excel file to analyze
     * @return An AnalysisResult containing statistical information about the file
     */
    AnalysisResult analyzeExcelFile(String filename);
}
