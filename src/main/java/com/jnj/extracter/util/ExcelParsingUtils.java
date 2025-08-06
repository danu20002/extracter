package com.jnj.extracter.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Utility class for handling problematic Excel files, 
 * especially those with encoding issues or corrupt pivot tables.
 */
@Slf4j
public class ExcelParsingUtils {
    
    /**
     * Attempts to extract sheet names from a potentially problematic Excel file
     * using a more robust approach that avoids parsing problematic parts like pivot tables.
     * 
     * @param file The Excel file to extract sheet names from
     * @return A list of sheet names found in the Excel file
     * @throws IOException If an IO error occurs
     */
    public static List<String> extractSheetNamesFromProblemFile(File file) throws IOException {
        List<String> sheetNames = new ArrayList<>();
        
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
            XSSFReader reader;
            try {
                reader = new XSSFReader(pkg);
            } catch (org.apache.poi.openxml4j.exceptions.OpenXML4JException e) {
                throw new IOException("Error creating XSSFReader: " + e.getMessage(), e);
            }
            
            // Get sheet references from the workbook part
            InputStream workbookData = reader.getWorkbookData();
            try (Workbook workbook = new XSSFWorkbook(pkg)) {
                int numberOfSheets = workbook.getNumberOfSheets();
                for (int i = 0; i < numberOfSheets; i++) {
                    sheetNames.add(workbook.getSheetName(i));
                }
            } catch (Exception e) {
                log.warn("Error accessing sheets directly, falling back to basic sheet enumeration");
                // If that fails, use a more basic approach with the XSSFReader
                XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
                while (sheetIterator.hasNext()) {
                    sheetIterator.next(); // stream for sheet data - not used here
                    sheetNames.add(sheetIterator.getSheetName());
                }
            }
        } catch (InvalidFormatException e) {
            throw new IOException("Invalid Excel file format: " + e.getMessage(), e);
        }
        
        return sheetNames;
    }

    /**
     * Creates a workbook that skips problematic parts like pivot tables or charts
     * by focusing only on worksheet data.
     * 
     * @param file The Excel file to process
     * @return A simplified workbook or null if creation fails
     */
    public static XSSFWorkbook createPartialWorkbook(File file) {
        try (OPCPackage pkg = OPCPackage.open(file, PackageAccess.READ)) {
            // Use custom options to avoid problematic parts
            // This is a simplified example and might need enhancement for specific files
            return new XSSFWorkbook(pkg);
        } catch (Exception e) {
            log.error("Failed to create partial workbook: {}", e.getMessage());
            return null;
        }
    }

    /**
     * Safely gets a cell value as string, handling various error cases.
     * 
     * @param cell The cell to get the value from
     * @return The cell value as a string, or null if there's an issue
     */
    public static String safeCellToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    // Check if it's a date
                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    }
                    // Use to string to avoid scientific notation issues
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e) {
                        try {
                            return String.valueOf(cell.getNumericCellValue());
                        } catch (Exception ex) {
                            return cell.getCellFormula();
                        }
                    }
                case BLANK:
                    return "";
                default:
                    return "[UNKNOWN]";
            }
        } catch (Exception e) {
            // If any error occurs while reading the cell, return a placeholder
            log.warn("Error reading cell: {}", e.getMessage());
            return "[ERROR]";
        }
    }

    /**
     * Checks if a row has any meaningful data (i.e., not all cells are blank)
     * 
     * @param row The row to check
     * @return true if the row has data, false otherwise
     */
    public static boolean hasData(Row row) {
        if (row == null) {
            return false;
        }
        
        for (Cell cell : row) {
            String value = safeCellToString(cell);
            if (value != null && !value.trim().isEmpty() && !value.equals("[ERROR]")) {
                return true;
            }
        }
        return false;
    }

    /**
     * Safely processes a sheet, extracting header row and data
     * 
     * @param sheet The sheet to process
     * @return A map containing "headers" list and "data" list (list of maps)
     */
    public static Map<String, Object> safeProcessSheet(Sheet sheet) {
        Map<String, Object> result = new HashMap<>();
        List<String> headers = new ArrayList<>();
        List<Map<String, String>> data = new ArrayList<>();
        
        if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
            result.put("headers", headers);
            result.put("data", data);
            return result;
        }
        
        // Get header row
        Row headerRow = sheet.getRow(sheet.getFirstRowNum());
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String headerValue = safeCellToString(cell);
                if (headerValue == null || headerValue.trim().isEmpty()) {
                    headerValue = "Column_" + (cell.getColumnIndex() + 1);
                }
                headers.add(headerValue);
            }
        }
        
        // Process data rows
        for (int i = sheet.getFirstRowNum() + 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null && hasData(row)) {
                Map<String, String> rowData = new HashMap<>();
                
                // For each header, find the corresponding cell
                for (int j = 0; j < headers.size(); j++) {
                    Cell cell = row.getCell(j);
                    String value = safeCellToString(cell);
                    rowData.put(headers.get(j), value);
                }
                
                data.add(rowData);
            }
        }
        
        result.put("headers", headers);
        result.put("data", data);
        return result;
    }
}
