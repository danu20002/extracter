package com.jnj.extracter.service.contract;

import io.micrometer.core.instrument.Timer;

import java.util.function.Supplier;

/**
 * Service interface for collecting and reporting metrics.
 */
public interface MetricsService {
    
    /**
     * Increment the counter for files processed.
     */
    void incrementFilesProcessed();
    
    /**
     * Increment the counter for rows processed.
     * 
     * @param count Number of rows processed
     */
    void incrementRowsProcessed(int count);
    
    /**
     * Increment the counter for processing errors.
     * 
     * @param errorType Type of error
     */
    void incrementProcessingErrors(String errorType);
    
    /**
     * Record the time taken to process a file.
     * 
     * @param fileName Name of the file
     * @param timeMs Processing time in milliseconds
     */
    void recordFileProcessingTime(String fileName, long timeMs);
    
    /**
     * Record the time taken to process a sheet.
     * 
     * @param sheetName Name of the sheet
     * @param timeMs Processing time in milliseconds
     */
    void recordSheetProcessingTime(String sheetName, long timeMs);
    
    /**
     * Record the execution time of a task using a specific timer.
     * 
     * @param <T> Return type of the task
     * @param timer The timer to use
     * @param task The task to execute
     * @return Result of the task
     */
    <T> T recordExecutionTime(Timer timer, Supplier<T> task);
    
    /**
     * Record the execution time of a task using a named timer.
     * 
     * @param <T> Return type of the task
     * @param timerName Name of the timer
     * @param task The task to execute
     * @param tags Optional tags as key-value pairs
     * @return Result of the task
     */
    <T> T recordExecutionTime(String timerName, Supplier<T> task, String... tags);
}
