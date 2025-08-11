package com.jnj.extracter.service.contract;

import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.ExcelFileInfo;
import com.jnj.extracter.domain.model.ExcelProcessingResult;
import com.jnj.extracter.domain.model.SheetData;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Service interface for Excel file operations.
 */
public interface ExcelService {
    
    /**
     * Retrieves information about all Excel files in the storage.
     *
     * @return List of ExcelFileInfo objects
     */
    List<ExcelFileInfo> getAllExcelFiles();
    
    /**
     * Retrieves all Excel files as File objects.
     *
     * @return List of File objects representing Excel files
     */
    List<File> getExcelFiles();
    
    /**
     * Gets the names of all sheets in the specified Excel file.
     *
     * @param file The Excel file
     * @return List of sheet names
     * @throws IOException If an error occurs reading the file
     */
    List<String> getSheetNames(File file) throws IOException;
    
    /**
     * Gets the Excel data for the specified file.
     *
     * @param filename The name of the file
     * @return The ExcelData object containing all data from the file
     */
    ExcelData getExcelData(String filename);
    
    /**
     * Gets data for a specific sheet in an Excel file.
     *
     * @param filename The name of the file
     * @param sheetName The name of the sheet
     * @return The SheetData object containing the sheet's data
     */
    SheetData getSheetData(String filename, String sheetName);
    
    /**
     * Processes an uploaded Excel file.
     *
     * @param file The MultipartFile containing the Excel file
     * @return ExcelProcessingResult with processing information
     */
    ExcelProcessingResult processExcelFile(MultipartFile file);
    
    /**
     * Clears the Excel data cache.
     */
    void clearCache();
    
    /**
     * Removes a specific file from the cache.
     *
     * @param filename The name of the file to remove
     */
    void removeFromCache(String filename);
}
