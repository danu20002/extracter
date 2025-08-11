package com.jnj.extracter.service.contract;

import com.jnj.extracter.domain.model.ExcelData;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;

/**
 * Service interface for reading Excel files.
 */
public interface ExcelReaderService {
    
    /**
     * Open an Excel workbook from a file.
     * 
     * @param file The Excel file
     * @return The opened Workbook
     */
    Workbook openWorkbook(File file);
    
    /**
     * Get sheet names from an Excel file.
     * 
     * @param file The Excel file
     * @return List of sheet names
     */
    List<String> getSheetNames(File file);
    
    /**
     * Extract value from a cell.
     * 
     * @param cell The Excel cell
     * @return The cell value as an Object
     */
    Object getCellValue(Cell cell);
    
    /**
     * Extract headers from a sheet.
     * 
     * @param sheet The Excel sheet
     * @return List of header values
     */
    List<String> extractHeaders(Sheet sheet);
    
    /**
     * Get the number of rows in a sheet.
     * 
     * @param sheet The Excel sheet
     * @return The row count
     */
    int getRowCount(Sheet sheet);
    
    /**
     * Extract row data from a sheet.
     * 
     * @param sheet The Excel sheet
     * @param rowIndex The row index
     * @param headers The headers list
     * @return Map of header name to cell value
     */
    Map<String, Object> extractRow(Sheet sheet, int rowIndex, List<String> headers);
    
    /**
     * Read an Excel file and convert it to ExcelData model.
     * 
     * @param filePath The path to the Excel file
     * @return The Excel data
     * @throws IOException If there's an error reading the file
     */
    ExcelData readExcelFile(Path filePath) throws IOException;
}
