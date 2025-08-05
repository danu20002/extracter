package com.jnj.extracter.service;

import com.jnj.extracter.entity.ExcelData;
import com.jnj.extracter.entity.ExcelProcessingResult;

import java.io.File;
import java.util.List;
import java.util.Map;

public interface ExcelService {
    
    /**
     * Extract data from all Excel files in the excel folder
     */
    List<ExcelProcessingResult> extractAllExcelFiles();
    
    /**
     * Extract data from a specific Excel file
     */
    ExcelProcessingResult extractExcelFile(File file);
    
    /**
     * Extract data from a specific sheet in an Excel file
     */
    List<ExcelData> extractSheetData(File file, String sheetName);
    
    /**
     * Get all available Excel files in the excel folder
     */
    List<File> getExcelFiles();
    
    /**
     * Perform operations on extracted data
     */
    Map<String, Object> performDataOperations(List<ExcelData> data, String operation);
    
    /**
     * Get summary statistics of the extracted data
     */
    Map<String, Object> getDataSummary(List<ExcelData> data);
}
