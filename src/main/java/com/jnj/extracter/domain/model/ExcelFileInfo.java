package com.jnj.extracter.domain.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Simple representation of an Excel file with relevant metadata
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ExcelFileInfo {
    /**
     * The name of the file
     */
    private String filename;
    
    /**
     * The size of the file in bytes
     */
    private long size;
    
    /**
     * The path to the file
     */
    private String path;
    
    /**
     * The last modified timestamp (formatted)
     */
    private String lastModified;
}
