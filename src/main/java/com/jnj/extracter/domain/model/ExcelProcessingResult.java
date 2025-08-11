package com.jnj.extracter.domain.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Represents the result of processing an Excel file.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelProcessingResult {
    /**
     * The name of the Excel file
     */
    private String filename;
    
    /**
     * Whether the processing was successful
     */
    private boolean success;
    
    /**
     * Processing message or error description
     */
    private String message;
    
    /**
     * Total number of sheets in the Excel file
     */
    private int sheetCount;
    
    /**
     * Total number of rows processed
     */
    private int totalRowCount;
    
    /**
     * Names of all sheets in the Excel file
     */
    private List<String> sheetNames;
    
    /**
     * The extracted data from the Excel file
     */
    private List<ExcelData> extractedData;
}
