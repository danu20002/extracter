package com.jnj.extracter.service.impl;

import com.jnj.extracter.api.exception.ExcelProcessingException;
import com.jnj.extracter.domain.model.ExcelRow;
import com.jnj.extracter.service.contract.TransformService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Implementation of TransformService for transforming Excel data.
 */
@Service
@Slf4j
public class TransformServiceImpl implements TransformService {

    private static final DateTimeFormatter TIMESTAMP_FORMATTER = 
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    @Override
    public List<ExcelRow> combineColumns(
            List<ExcelRow> data,
            List<String> sourceColumns,
            String targetColumn,
            String separator) {
        
        log.info("Combining columns {} into {} using separator '{}'", 
                sourceColumns, targetColumn, separator);
        
        if (separator == null) {
            separator = "";
        }
        
        List<ExcelRow> transformedData = new ArrayList<>(data.size());
        
        for (ExcelRow row : data) {
            Map<String, Object> newData = new HashMap<>(row.getData());
            StringBuilder combined = new StringBuilder();
            
            // Combine values from source columns
            for (int i = 0; i < sourceColumns.size(); i++) {
                String column = sourceColumns.get(i);
                Object value = row.getData().get(column);
                
                if (value != null) {
                    if (i > 0 && combined.length() > 0) {
                        combined.append(separator);
                    }
                    combined.append(value.toString());
                }
            }
            
            // Add the new column
            newData.put(targetColumn, combined.toString());
            
            // Create a new ExcelRow instance with the transformed data
            transformedData.add(ExcelRow.builder()
                .fileName(row.getFileName())
                .sheetName(row.getSheetName())
                .rowNumber(row.getRowNumber())
                .data(newData)
                .extractedAt(row.getExtractedAt())
                .build());
        }
        
        log.info("Successfully combined columns for {} rows", transformedData.size());
        return transformedData;
    }

    @Override
    public String createTransformedExcelFile(
            List<?> data,
            Map<String, List<String>> transformationMap,
            Map<String, String> separatorMap,
            String outputFileName,
            boolean includeOriginalColumns) {
        
        log.info("Creating transformed Excel file: {}", outputFileName);
        
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Transformed Data");
            
            // Apply all transformations to the data
            List<ExcelRow> transformedData;
            
            // Convert the input data to List<ExcelRow> if needed
            if (data.isEmpty()) {
                transformedData = new ArrayList<>();
            } else if (data.get(0) instanceof ExcelRow) {
                @SuppressWarnings("unchecked")
                List<ExcelRow> castedData = (List<ExcelRow>) data;
                transformedData = castedData;
            } else {
                // Convert generic data to ExcelRow objects
                transformedData = new ArrayList<>();
                for (Object item : data) {
                    if (item instanceof Map) {
                        @SuppressWarnings("unchecked")
                        Map<String, Object> map = (Map<String, Object>) item;
                        transformedData.add(ExcelRow.builder()
                            .fileName("generated")
                            .sheetName("Sheet1")
                            .rowNumber(transformedData.size())
                            .data(new HashMap<>(map))
                            .extractedAt(LocalDateTime.now())
                            .build());
                    }
                }
            }
            
            for (Map.Entry<String, List<String>> transformation : transformationMap.entrySet()) {
                String targetColumn = transformation.getKey();
                List<String> sourceColumns = transformation.getValue();
                String separator = separatorMap != null ? separatorMap.get(targetColumn) : null;
                
                transformedData = combineColumns(transformedData, sourceColumns, targetColumn, separator);
            }
            
            // Get all column names
            Set<String> columnNames = new HashSet<>();
            
            // Add transformed columns
            columnNames.addAll(transformationMap.keySet());
            
            // Add original columns if required
            if (includeOriginalColumns && !transformedData.isEmpty()) {
                columnNames.addAll(transformedData.get(0).getData().keySet());
            }
            
            // Remove source columns if they should not be included
            if (!includeOriginalColumns) {
                for (List<String> sourceColumns : transformationMap.values()) {
                    columnNames.removeAll(sourceColumns);
                }
            }
            
            List<String> sortedColumns = new ArrayList<>(columnNames);
            Collections.sort(sortedColumns);
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < sortedColumns.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(sortedColumns.get(i));
            }
            
            // Create data rows
            for (int i = 0; i < transformedData.size(); i++) {
                Row row = sheet.createRow(i + 1);
                ExcelRow rowData = transformedData.get(i);
                
                for (int j = 0; j < sortedColumns.size(); j++) {
                    Cell cell = row.createCell(j);
                    Object value = rowData.getData().get(sortedColumns.get(j));
                    
                    if (value != null) {
                        if (value instanceof Number) {
                            cell.setCellValue(((Number) value).doubleValue());
                        } else if (value instanceof Date) {
                            cell.setCellValue((Date) value);
                        } else if (value instanceof Boolean) {
                            cell.setCellValue((Boolean) value);
                        } else {
                            cell.setCellValue(value.toString());
                        }
                    }
                }
            }
            
            // Auto-size columns
            for (int i = 0; i < sortedColumns.size(); i++) {
                sheet.autoSizeColumn(i);
            }
            
            // Write to file
            String filePath = "excel/temp/" + outputFileName;
            if (!outputFileName.endsWith(".xlsx")) {
                filePath += ".xlsx";
            }
            
            File outputFile = new File(filePath);
            outputFile.getParentFile().mkdirs();
            
            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                workbook.write(fileOut);
            }
            
            log.info("Successfully created transformed Excel file: {}", filePath);
            return filePath;
            
        } catch (IOException e) {
            log.error("Failed to create transformed Excel file", e);
            throw new ExcelProcessingException("Failed to create transformed Excel file: " + e.getMessage(), e);
        }
    }

    @Override
    public List<ExcelRow> splitColumn(
            List<ExcelRow> data,
            String sourceColumn,
            List<String> targetColumns,
            String delimiter) {
        
        log.info("Splitting column {} into {} using delimiter '{}'",
                sourceColumn, targetColumns, delimiter);
        
        if (delimiter == null || delimiter.isEmpty()) {
            throw new IllegalArgumentException("Delimiter cannot be null or empty for splitting");
        }
        
        List<ExcelRow> transformedData = new ArrayList<>(data.size());
        
        for (ExcelRow row : data) {
            Map<String, Object> newData = new HashMap<>(row.getData());
            Object sourceValue = row.getData().get(sourceColumn);
            
            if (sourceValue != null) {
                String[] parts = sourceValue.toString().split(delimiter);
                
                for (int i = 0; i < targetColumns.size() && i < parts.length; i++) {
                    newData.put(targetColumns.get(i), parts[i].trim());
                }
                
                // Fill remaining target columns with empty strings
                for (int i = parts.length; i < targetColumns.size(); i++) {
                    newData.put(targetColumns.get(i), "");
                }
            } else {
                // If source value is null, set all target columns to empty
                for (String targetColumn : targetColumns) {
                    newData.put(targetColumn, "");
                }
            }
            
            transformedData.add(ExcelRow.builder()
                .fileName(row.getFileName())
                .sheetName(row.getSheetName())
                .rowNumber(row.getRowNumber())
                .data(newData)
                .extractedAt(row.getExtractedAt())
                .build());
        }
        
        log.info("Successfully split column for {} rows", transformedData.size());
        return transformedData;
    }

    @Override
    public List<ExcelRow> transformColumn(
            List<ExcelRow> data,
            String sourceColumn,
            String targetColumn,
            String transformationType) {
        
        log.info("Transforming column {} to {} using transformation type: {}",
                sourceColumn, targetColumn, transformationType);
        
        List<ExcelRow> transformedData = new ArrayList<>(data.size());
        
        for (ExcelRow row : data) {
            Map<String, Object> newData = new HashMap<>(row.getData());
            Object sourceValue = row.getData().get(sourceColumn);
            
            if (sourceValue != null) {
                Object transformedValue = applyTransformation(sourceValue, transformationType);
                newData.put(targetColumn, transformedValue);
            }
            
            transformedData.add(ExcelRow.builder()
                .fileName(row.getFileName())
                .sheetName(row.getSheetName())
                .rowNumber(row.getRowNumber())
                .data(newData)
                .extractedAt(row.getExtractedAt())
                .build());
        }
        
        log.info("Successfully transformed column for {} rows", transformedData.size());
        return transformedData;
    }
    
    /**
     * Apply a transformation to a value.
     * 
     * @param value The value to transform
     * @param transformationType The type of transformation to apply
     * @return The transformed value
     */
    private Object applyTransformation(Object value, String transformationType) {
        if (value == null) {
            return null;
        }
        
        String stringValue = value.toString();
        
        switch (transformationType.toLowerCase()) {
            case "uppercase":
                return stringValue.toUpperCase();
            case "lowercase":
                return stringValue.toLowerCase();
            case "trim":
                return stringValue.trim();
            case "number":
                try {
                    return Double.parseDouble(stringValue);
                } catch (NumberFormatException e) {
                    return value;
                }
            case "date":
                // Simple date parsing - would need more sophisticated logic in real implementation
                try {
                    return LocalDateTime.parse(stringValue, TIMESTAMP_FORMATTER);
                } catch (Exception e) {
                    return value;
                }
            default:
                return value;
        }
    }
}
