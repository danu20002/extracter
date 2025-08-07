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
     * Get sheet names from a specific Excel file
     */
    List<String> getSheetNames(File file);
    
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
    
    /**
     * Transform data by combining columns
     * 
     * @param data The Excel data to transform
     * @param sourceColumns List of source column names to combine
     * @param targetColumn Name of the new column to create
     * @param separator Optional separator between combined values (default: no separator)
     * @return The transformed Excel data
     */
    List<ExcelData> transformDataByCombiningColumns(List<ExcelData> data, List<String> sourceColumns, 
                                                  String targetColumn, String separator);
                                                  
    /**
     * Create a new Excel file with multiple transformed columns
     * 
     * @param data The Excel data to transform
     * @param transformationMap Map of target column names to lists of source columns
     * @param separatorMap Map of target column names to separators (optional)
     * @param outputFileName Name of the output Excel file
     * @param includeOriginalColumns Whether to include original columns in the output
     * @return Path to the created Excel file
     */
    String createTransformedExcelFile(List<ExcelData> data, 
                                     Map<String, List<String>> transformationMap,
                                     Map<String, String> separatorMap,
                                     String outputFileName,
                                     boolean includeOriginalColumns);
}
