package com.jnj.extracter.service.impl;

import com.jnj.extracter.api.exception.ExcelProcessingException;
import com.jnj.extracter.domain.model.CellData;
import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.RowData;
import com.jnj.extracter.domain.model.SheetData;
import com.jnj.extracter.service.contract.ExcelReaderService;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.*;

/**
 * Implementation of ExcelReaderService for reading Excel files.
 */
@Service
@Slf4j
public class ExcelReaderServiceImpl implements ExcelReaderService {

    @Override
    public Workbook openWorkbook(File file) {
        try {
            log.debug("Opening workbook for file: {}", file.getName());
            return WorkbookFactory.create(new FileInputStream(file));
        } catch (IOException e) {
            log.error("Error opening Excel workbook: {}", file.getName(), e);
            throw new ExcelProcessingException("Failed to open Excel file: " + file.getName(), e);
        }
    }

    @Override
    public List<String> getSheetNames(File file) {
        try (Workbook workbook = openWorkbook(file)) {
            log.debug("Getting sheet names for file: {}", file.getName());
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheetNames.add(workbook.getSheetName(i));
            }
            return sheetNames;
        } catch (IOException e) {
            log.error("Error getting sheet names: {}", file.getName(), e);
            throw new ExcelProcessingException("Failed to get sheet names: " + file.getName(), e);
        }
    }

    @Override
    public Object getCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                // Handle date cells
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                try {
                    // Try to evaluate formula
                    return cell.getNumericCellValue();
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            case ERROR:
                return "ERROR: " + cell.getErrorCellValue();
            default:
                return null;
        }
    }

    @Override
    public List<String> extractHeaders(Sheet sheet) {
        log.debug("Extracting headers from sheet: {}", sheet.getSheetName());
        Row headerRow = sheet.getRow(0);
        List<String> headers = new ArrayList<>();

        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String headerValue = String.valueOf(getCellValue(cell));
                headers.add(headerValue);
            }
        }

        return headers;
    }

    @Override
    public int getRowCount(Sheet sheet) {
        return sheet.getLastRowNum() + 1;
    }

    @Override
    public Map<String, Object> extractRow(Sheet sheet, int rowIndex, List<String> headers) {
        Row row = sheet.getRow(rowIndex);
        Map<String, Object> rowData = new HashMap<>();

        if (row != null) {
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                rowData.put(headers.get(i), getCellValue(cell));
            }
        }

        return rowData;
    }

    @Override
    public ExcelData readExcelFile(Path filePath) throws IOException {
        log.info("Reading Excel file: {}", filePath.getFileName());
        File file = filePath.toFile();

        try (Workbook workbook = openWorkbook(file)) {
            String filename = filePath.getFileName().toString();

            // Create Excel data structure
            ExcelData excelData = new ExcelData();
            excelData.setFilename(filename);
            excelData.setSheetCount(workbook.getNumberOfSheets());
            
            List<SheetData> sheets = new ArrayList<>();

            // Process each sheet
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                String sheetName = sheet.getSheetName();
                log.debug("Processing sheet: {}", sheetName);

                // Create sheet data
                SheetData sheetData = new SheetData();
                sheetData.setName(sheetName);
                sheetData.setIndex(i);

                // Extract headers
                List<String> headers = extractHeaders(sheet);
                sheetData.setHeaders(headers);

                // Extract rows
                List<RowData> rows = new ArrayList<>();
                // Start from row 1 to skip header
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) {
                        continue;
                    }

                    RowData rowData = new RowData();
                    rowData.setIndex(rowIndex);
                    rowData.setHeader(false);

                    List<CellData> cells = new ArrayList<>();
                    for (int colIndex = 0; colIndex < headers.size(); colIndex++) {
                        Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        
                        CellData cellData = new CellData();
                        cellData.setRowIndex(rowIndex);
                        cellData.setColumnIndex(colIndex);
                        cellData.setHeader(colIndex < headers.size() ? headers.get(colIndex) : "Column " + colIndex);
                        
                        Object value = getCellValue(cell);
                        cellData.setValue(value != null ? value.toString() : "");
                        cellData.setType(cell.getCellType().toString());
                        
                        if (cell.getCellType() == CellType.FORMULA) {
                            cellData.setFormula(cell.getCellFormula());
                        }
                        
                        cells.add(cellData);
                    }
                    
                    rowData.setCells(cells);
                    rows.add(rowData);
                }

                sheetData.setRows(rows);
                sheets.add(sheetData);
            }

            excelData.setSheets(sheets);
            log.info("Completed reading Excel file: {} with {} sheets", filename, sheets.size());
            return excelData;
        } catch (Exception e) {
            log.error("Error processing Excel file: {}", filePath, e);
            throw new ExcelProcessingException("Failed to process Excel file: " + filePath, e);
        }
    }
}
