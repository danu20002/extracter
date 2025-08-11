package com.jnj.extracter.domain.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

/**
 * Represents Excel data including all sheets.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelData {
    /**
     * The name of the Excel file
     */
    private String filename;
    
    /**
     * The number of sheets in the Excel file
     */
    private int sheetCount;
    
    /**
     * List of sheet data
     */
    private List<SheetData> sheets;
}
