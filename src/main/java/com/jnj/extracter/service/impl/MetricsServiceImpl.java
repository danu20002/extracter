package com.jnj.extracter.service.impl;

import com.jnj.extracter.service.contract.MetricsService;
import io.micrometer.core.instrument.Counter;
import io.micrometer.core.instrument.MeterRegistry;
import io.micrometer.core.instrument.Tag;
import io.micrometer.core.instrument.Timer;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.Arrays;
import java.util.concurrent.TimeUnit;
import java.util.function.Supplier;

/**
 * Implementation of MetricsService for collecting and reporting metrics about Excel processing.
 */
@Service
@Slf4j
public class MetricsServiceImpl implements MetricsService {

    private final MeterRegistry registry;
    
    // Counters
    private final Counter filesProcessedCounter;
    private final Counter rowsProcessedCounter;
    private final Counter processingErrorsCounter;
    
    // Timers
    private final Timer fileProcessingTimer;
    private final Timer sheetProcessingTimer;
    
    public MetricsServiceImpl(MeterRegistry registry) {
        this.registry = registry;
        
        // Initialize counters
        this.filesProcessedCounter = registry.counter("excel.files.processed");
        this.rowsProcessedCounter = registry.counter("excel.rows.processed");
        this.processingErrorsCounter = registry.counter("excel.processing.errors");
        
        // Initialize timers
        this.fileProcessingTimer = registry.timer("excel.file.processing");
        this.sheetProcessingTimer = registry.timer("excel.sheet.processing");
        
        log.info("MetricsService initialized with registry: {}", registry.getClass().getSimpleName());
    }
    
    @Override
    public void incrementFilesProcessed() {
        filesProcessedCounter.increment();
        log.debug("Files processed counter incremented");
    }
    
    @Override
    public void incrementRowsProcessed(int count) {
        rowsProcessedCounter.increment(count);
        log.debug("Rows processed counter incremented by {}", count);
    }
    
    @Override
    public void incrementProcessingErrors(String errorType) {
        processingErrorsCounter.increment();
        registry.counter("excel.processing.errors.by.type", 
                Arrays.asList(Tag.of("errorType", errorType))).increment();
        log.debug("Processing errors counter incremented for type: {}", errorType);
    }
    
    @Override
    public void recordFileProcessingTime(String fileName, long timeMs) {
        fileProcessingTimer.record(timeMs, TimeUnit.MILLISECONDS);
        
        // Also record with file name tag for more detailed metrics
        registry.timer("excel.file.processing.detailed", 
                Arrays.asList(Tag.of("fileName", fileName)))
                .record(timeMs, TimeUnit.MILLISECONDS);
                
        log.debug("File processing time recorded: {} ms for {}", timeMs, fileName);
    }
    
    @Override
    public void recordSheetProcessingTime(String sheetName, long timeMs) {
        sheetProcessingTimer.record(timeMs, TimeUnit.MILLISECONDS);
        
        // Also record with sheet name tag for more detailed metrics
        registry.timer("excel.sheet.processing.detailed", 
                Arrays.asList(Tag.of("sheetName", sheetName)))
                .record(timeMs, TimeUnit.MILLISECONDS);
                
        log.debug("Sheet processing time recorded: {} ms for {}", timeMs, sheetName);
    }
    
    @Override
    public <T> T recordExecutionTime(Timer timer, Supplier<T> task) {
        return timer.record(task);
    }
    
    @Override
    public <T> T recordExecutionTime(String timerName, Supplier<T> task, String... tags) {
        if (tags.length % 2 != 0) {
            throw new IllegalArgumentException("Tags must be provided as key-value pairs");
        }
        
        Timer.Sample sample = Timer.start(registry);
        try {
            T result = task.get();
            return result;
        } finally {
            if (tags.length > 0) {
                Tag[] tagArray = new Tag[tags.length / 2];
                for (int i = 0; i < tags.length; i += 2) {
                    tagArray[i/2] = Tag.of(tags[i], tags[i+1]);
                }
                sample.stop(registry.timer(timerName, Arrays.asList(tagArray)));
            } else {
                sample.stop(registry.timer(timerName));
            }
        }
    }
}
