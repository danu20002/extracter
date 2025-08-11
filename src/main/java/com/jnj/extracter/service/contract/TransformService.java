package com.jnj.extracter.service.contract;

import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.ExcelRow;

import java.util.List;
import java.util.Map;

/**
 * Service interface for transforming Excel data.
 * This interface separates data transformation functionality 
 * from other Excel processing operations.
 */
public interface TransformService {
    
    /**
     * Transform Excel data by combining columns.
     * 
     * @param data The Excel rows to transform
     * @param sourceColumns List of source column names to combine
     * @param targetColumn Name of the new column to create
     * @param separator Separator between combined values (default: no separator)
     * @return The transformed Excel rows
     */
    List<ExcelRow> combineColumns(
            List<ExcelRow> data, 
            List<String> sourceColumns, 
            String targetColumn, 
            String separator);
    
    /**
     * Create a new Excel file with multiple transformed columns.
     * 
     * @param data The Excel data to transform
     * @param transformationMap Map of target column names to lists of source columns
     * @param separatorMap Map of target column names to separators
     * @param outputFileName Name of the output Excel file
     * @param includeOriginalColumns Whether to include original columns in the output
     * @return Path to the created Excel file
     */
    String createTransformedExcelFile(
            List<?> data, 
            Map<String, List<String>> transformationMap,
            Map<String, String> separatorMap,
            String outputFileName,
            boolean includeOriginalColumns);
    
    /**
     * Split a column into multiple columns.
     * 
     * @param data The Excel rows
     * @param sourceColumn The column to split
     * @param targetColumns Names for the new columns
     * @param delimiter The delimiter to split by
     * @return The transformed Excel rows
     */
    List<ExcelRow> splitColumn(
            List<ExcelRow> data, 
            String sourceColumn, 
            List<String> targetColumns, 
            String delimiter);
    
    /**
     * Apply a transformation function to a column.
     * 
     * @param data The Excel rows
     * @param sourceColumn The column to transform
     * @param targetColumn The new column name (can be same as sourceColumn to overwrite)
     * @param transformationType Type of transformation to apply
     * @return The transformed Excel rows
     */
    List<ExcelRow> transformColumn(
            List<ExcelRow> data, 
            String sourceColumn, 
            String targetColumn, 
            String transformationType);
}
