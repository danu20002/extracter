package com.jnj.extracter.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelData {
    private String fileName;
    private String sheetName;
    private int rowNumber;
    private Map<String, Object> data;
    private String extractedAt;
}
