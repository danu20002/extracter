package com.jnj.extracter.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelProcessingResult {
    private String fileName;
    private boolean success;
    private String message;
    private int totalSheets;
    private int totalRows;
    private List<String> sheetNames;
    private List<ExcelData> extractedData;
}
