package com.jnj.extracter.service.impl;

import com.jnj.extracter.api.exception.ExcelProcessingException;
import com.jnj.extracter.domain.model.AnalysisResult;
import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.SheetData;
import com.jnj.extracter.service.contract.AnalysisService;
import com.jnj.extracter.service.contract.ExcelService;
import com.jnj.extracter.service.contract.MetricsService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Implementation of AnalysisService that performs statistical analysis on Excel data.
 */
@Service
@Slf4j
public class AnalysisServiceImpl implements AnalysisService {

    private final ExcelService excelService;
    private final MetricsService metricsService;
    
    public AnalysisServiceImpl(ExcelService excelService, MetricsService metricsService) {
        this.excelService = excelService;
        this.metricsService = metricsService;
        log.info("AnalysisServiceImpl initialized");
    }
    
    @Override
    public AnalysisResult analyzeExcelFile(String filename) {
        log.debug("Analyzing Excel file: {}", filename);
        
        try {
            long startTime = System.currentTimeMillis();
            
            ExcelData excelData = excelService.getExcelData(filename);
            if (excelData == null || excelData.getSheets() == null || excelData.getSheets().isEmpty()) {
                throw new ExcelProcessingException("No data found in file: " + filename);
            }
            
            AnalysisResult result = new AnalysisResult();
            result.setFilename(filename);
            result.setSheetCount(excelData.getSheets().size());
            
            // File statistics
            Map<String, Object> fileStats = calculateFileStatistics(excelData);
            result.setFileStatistics(fileStats);
            
            // Sheet statistics
            Map<String, Map<String, Object>> sheetStats = calculateSheetStatistics(excelData.getSheets());
            result.setSheetStatistics(sheetStats);
            
            // Column statistics
            Map<String, Map<String, Map<String, Object>>> columnStats = calculateColumnStatistics(excelData.getSheets());
            result.setColumnStatistics(columnStats);
            
            long processingTime = System.currentTimeMillis() - startTime;
            log.info("Analysis completed for file {} in {}ms", filename, processingTime);
            
            return result;
            
        } catch (Exception e) {
            log.error("Error analyzing Excel file: {}", filename, e);
            metricsService.incrementProcessingErrors("analysis_error");
            throw new ExcelProcessingException("Failed to analyze Excel file: " + filename, e);
        }
    }
    
    private Map<String, Object> calculateFileStatistics(ExcelData excelData) {
        Map<String, Object> stats = new HashMap<>();
        
        int totalRows = excelData.getSheets().stream()
                .mapToInt(sheet -> sheet.getRows().size())
                .sum();
        
        int totalCells = excelData.getSheets().stream()
                .flatMap(sheet -> sheet.getRows().stream())
                .mapToInt(row -> row.getCells().size())
                .sum();
        
        double averageRowsPerSheet = excelData.getSheets().stream()
                .mapToInt(sheet -> sheet.getRows().size())
                .average()
                .orElse(0);
        
        stats.put("totalSheets", excelData.getSheets().size());
        stats.put("totalRows", totalRows);
        stats.put("totalCells", totalCells);
        stats.put("averageRowsPerSheet", Math.round(averageRowsPerSheet * 100.0) / 100.0);
        
        return stats;
    }
    
    private Map<String, Map<String, Object>> calculateSheetStatistics(List<SheetData> sheets) {
        Map<String, Map<String, Object>> sheetStats = new HashMap<>();
        
        for (SheetData sheet : sheets) {
            Map<String, Object> stats = new HashMap<>();
            
            int rowCount = sheet.getRows().size();
            int headerCount = sheet.getHeaders() != null ? sheet.getHeaders().size() : 0;
            int emptyCellCount = (int) sheet.getRows().stream()
                    .flatMap(row -> row.getCells().stream())
                    .filter(cell -> cell.getValue() == null || cell.getValue().isEmpty())
                    .count();
            
            int totalCells = sheet.getRows().stream()
                    .mapToInt(row -> row.getCells().size())
                    .sum();
            
            double emptyCellPercentage = totalCells > 0 ? 
                    (double) emptyCellCount / totalCells * 100 : 0;
            
            stats.put("rowCount", rowCount);
            stats.put("headerCount", headerCount);
            stats.put("emptyCellCount", emptyCellCount);
            stats.put("totalCellCount", totalCells);
            stats.put("emptyCellPercentage", Math.round(emptyCellPercentage * 100.0) / 100.0);
            
            sheetStats.put(sheet.getName(), stats);
        }
        
        return sheetStats;
    }
    
    private Map<String, Map<String, Map<String, Object>>> calculateColumnStatistics(List<SheetData> sheets) {
        Map<String, Map<String, Map<String, Object>>> columnStats = new HashMap<>();
        
        for (SheetData sheet : sheets) {
            Map<String, Map<String, Object>> sheetColumnStats = new HashMap<>();
            List<String> headers = sheet.getHeaders();
            
            if (headers != null && !headers.isEmpty()) {
                for (String header : headers) {
                    Map<String, Object> stats = analyzeColumn(sheet, header);
                    sheetColumnStats.put(header, stats);
                }
            }
            
            columnStats.put(sheet.getName(), sheetColumnStats);
        }
        
        return columnStats;
    }
    
    private Map<String, Object> analyzeColumn(SheetData sheet, String header) {
        Map<String, Object> stats = new HashMap<>();
        
        // Get all values for this column
        List<String> columnValues = sheet.getRows().stream()
                .flatMap(row -> row.getCells().stream())
                .filter(cell -> header.equals(cell.getHeader()))
                .map(cell -> cell.getValue())
                .filter(value -> value != null && !value.isEmpty())
                .collect(Collectors.toList());
        
        int totalValues = columnValues.size();
        int distinctValues = (int) columnValues.stream().distinct().count();
        int emptyValues = (int) sheet.getRows().stream()
                .flatMap(row -> row.getCells().stream())
                .filter(cell -> header.equals(cell.getHeader()))
                .filter(cell -> cell.getValue() == null || cell.getValue().isEmpty())
                .count();
        
        stats.put("totalValues", totalValues);
        stats.put("distinctValues", distinctValues);
        stats.put("emptyValues", emptyValues);
        
        // Determine data type
        String dataType = determineColumnDataType(columnValues);
        stats.put("inferredDataType", dataType);
        
        // If numeric, calculate statistics
        if ("NUMERIC".equals(dataType) && !columnValues.isEmpty()) {
            try {
                List<Double> numericValues = columnValues.stream()
                        .map(v -> {
                            try {
                                return Double.parseDouble(v);
                            } catch (NumberFormatException e) {
                                return null;
                            }
                        })
                        .filter(v -> v != null)
                        .collect(Collectors.toList());
                
                if (!numericValues.isEmpty()) {
                    double min = numericValues.stream().mapToDouble(v -> v).min().orElse(0);
                    double max = numericValues.stream().mapToDouble(v -> v).max().orElse(0);
                    double avg = numericValues.stream().mapToDouble(v -> v).average().orElse(0);
                    
                    stats.put("min", min);
                    stats.put("max", max);
                    stats.put("average", Math.round(avg * 100.0) / 100.0);
                }
            } catch (Exception e) {
                log.warn("Error calculating numeric statistics for column {}: {}", header, e.getMessage());
            }
        }
        
        return stats;
    }
    
    private String determineColumnDataType(List<String> values) {
        if (values == null || values.isEmpty()) {
            return "UNKNOWN";
        }
        
        boolean allNumeric = values.stream()
                .allMatch(v -> {
                    try {
                        Double.parseDouble(v);
                        return true;
                    } catch (NumberFormatException e) {
                        return false;
                    }
                });
        
        if (allNumeric) {
            return "NUMERIC";
        }
        
        boolean allDates = values.stream()
                .allMatch(v -> {
                    try {
                        // Simple date check - could be improved
                        return v.matches("\\d{4}-\\d{2}-\\d{2}.*") || 
                               v.matches("\\d{2}/\\d{2}/\\d{4}.*");
                    } catch (Exception e) {
                        return false;
                    }
                });
        
        if (allDates) {
            return "DATE";
        }
        
        boolean allBoolean = values.stream()
                .allMatch(v -> v.equalsIgnoreCase("true") || v.equalsIgnoreCase("false") || 
                               v.equals("0") || v.equals("1") || 
                               v.equalsIgnoreCase("yes") || v.equalsIgnoreCase("no"));
        
        if (allBoolean) {
            return "BOOLEAN";
        }
        
        return "STRING";
    }
}
