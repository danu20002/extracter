package com.jnj.extracter.service;

import io.micrometer.core.instrument.Counter;
import io.micrometer.core.instrument.MeterRegistry;
import io.micrometer.core.instrument.Tag;
import io.micrometer.core.instrument.Timer;
import org.springframework.stereotype.Service;

import java.util.Arrays;
import java.util.concurrent.TimeUnit;
import java.util.function.Supplier;

/**
 * Service for collecting and reporting metrics about Excel processing.
 */
@Service
public class MetricsService {

    private final MeterRegistry registry;
    
    // Counters
    private final Counter filesProcessedCounter;
    private final Counter rowsProcessedCounter;
    private final Counter processingErrorsCounter;
    
    // Timers
    private final Timer fileProcessingTimer;
    private final Timer sheetProcessingTimer;
    
    public MetricsService(MeterRegistry registry) {
        this.registry = registry;
        
        // Initialize counters
        this.filesProcessedCounter = registry.counter("excel.files.processed");
        this.rowsProcessedCounter = registry.counter("excel.rows.processed");
        this.processingErrorsCounter = registry.counter("excel.processing.errors");
        
        // Initialize timers
        this.fileProcessingTimer = registry.timer("excel.file.processing.time");
        this.sheetProcessingTimer = registry.timer("excel.sheet.processing.time");
    }
    
    /**
     * Record that a file has been processed.
     * 
     * @param fileType The type of file (extension)
     */
    public void recordFileProcessed(String fileType) {
        filesProcessedCounter.increment();
        registry.counter("excel.files.processed.by.type", Arrays.asList(Tag.of("fileType", fileType))).increment();
    }
    
    /**
     * Record that rows have been processed.
     * 
     * @param count The number of rows processed
     */
    public void recordRowsProcessed(int count) {
        rowsProcessedCounter.increment(count);
    }
    
    /**
     * Record an error that occurred during processing.
     * 
     * @param errorType The type of error
     */
    public void recordProcessingError(String errorType) {
        processingErrorsCounter.increment();
        registry.counter("excel.processing.errors.by.type", Arrays.asList(Tag.of("errorType", errorType))).increment();
    }
    
    /**
     * Record the time taken to process a file.
     * 
     * @param fileName The name of the file
     * @param timeMs The time taken in milliseconds
     */
    public void recordFileProcessingTime(String fileName, long timeMs) {
        fileProcessingTimer.record(timeMs, TimeUnit.MILLISECONDS);
        registry.timer("excel.file.processing.time.by.file", 
                      Arrays.asList(Tag.of("fileName", fileName)))
                .record(timeMs, TimeUnit.MILLISECONDS);
    }
    
    /**
     * Record the time taken to process a sheet.
     * 
     * @param sheetName The name of the sheet
     * @param timeMs The time taken in milliseconds
     */
    public void recordSheetProcessingTime(String sheetName, long timeMs) {
        sheetProcessingTimer.record(timeMs, TimeUnit.MILLISECONDS);
        registry.timer("excel.sheet.processing.time.by.sheet", 
                      Arrays.asList(Tag.of("sheetName", sheetName)))
                .record(timeMs, TimeUnit.MILLISECONDS);
    }
    
    /**
     * Execute an operation and record its execution time.
     * 
     * @param <T> The return type of the operation
     * @param timerName The name of the timer
     * @param tags Tags to associate with the timer
     * @param operation The operation to execute
     * @return The result of the operation
     */
    public <T> T recordExecutionTime(String timerName, Iterable<Tag> tags, Supplier<T> operation) {
        long startTime = System.currentTimeMillis();
        T result = operation.get();
        long endTime = System.currentTimeMillis();
        
        registry.timer(timerName, tags).record(endTime - startTime, TimeUnit.MILLISECONDS);
        
        return result;
    }
    
    /**
     * Record memory usage statistics.
     */
    public void recordMemoryUsage() {
        Runtime runtime = Runtime.getRuntime();
        long totalMemory = runtime.totalMemory();
        long freeMemory = runtime.freeMemory();
        long maxMemory = runtime.maxMemory();
        long usedMemory = totalMemory - freeMemory;
        
        registry.gauge("jvm.memory.used", usedMemory / 1024 / 1024); // MB
        registry.gauge("jvm.memory.free", freeMemory / 1024 / 1024); // MB
        registry.gauge("jvm.memory.total", totalMemory / 1024 / 1024); // MB
        registry.gauge("jvm.memory.max", maxMemory / 1024 / 1024); // MB
        registry.gauge("jvm.memory.used.percentage", (double) usedMemory / totalMemory * 100);
    }
}
