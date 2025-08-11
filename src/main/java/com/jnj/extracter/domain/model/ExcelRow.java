package com.jnj.extracter.domain.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;
import java.util.Map;

/**
 * Represents a row of data from an Excel file with additional metadata.
 * This class is used primarily for transformation operations.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelRow {
    /**
     * The name of the source Excel file
     */
    private String fileName;
    
    /**
     * The name of the sheet containing this row
     */
    private String sheetName;
    
    /**
     * The row number in the original sheet (0-based)
     */
    private int rowNumber;
    
    /**
     * The data in this row, as column name -> value mapping
     */
    private Map<String, Object> data;
    
    /**
     * When this data was extracted
     */
    private LocalDateTime extractedAt;
}
