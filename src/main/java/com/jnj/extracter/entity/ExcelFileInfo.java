package com.jnj.extracter.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Simple representation of an Excel file with relevant metadata
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelFileInfo {
    private String name;
    private long size;
    private String path;
    private String lastModified;
}
