package com.jnj.extracter.api.controller;

import io.micrometer.core.instrument.MeterRegistry;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.HashMap;
import java.util.Map;

/**
 * REST controller that exposes Excel processing statistics.
 * Provides metrics about Excel file processing operations.
 */
@RestController
@RequestMapping("/api/stats")
public class ExcelStatsController {

    private final MeterRegistry meterRegistry;
    
    public ExcelStatsController(MeterRegistry meterRegistry) {
        this.meterRegistry = meterRegistry;
    }
    
    /**
     * Get comprehensive Excel processing statistics.
     * 
     * @return Map containing various Excel processing metrics
     */
    @GetMapping
    public Map<String, Object> getExcelStats() {
        Map<String, Object> stats = new HashMap<>();
        
        // Get counters
        stats.put("filesProcessed", getCounterValue("excel.files.processed"));
        stats.put("rowsProcessed", getCounterValue("excel.rows.processed"));
        stats.put("errors", getCounterValue("excel.processing.errors"));
        
        // Get timer statistics
        Map<String, Object> timers = new HashMap<>();
        timers.put("fileProcessing", getTimerStats("excel.file.processing"));
        timers.put("sheetProcessing", getTimerStats("excel.sheet.processing"));
        
        stats.put("timers", timers);
        
        // Get gauge values
        stats.put("activeProcessingJobs", getGaugeValue("excel.active.jobs"));
        stats.put("memoryUsage", getMemoryStats());
        
        return stats;
    }
    
    /**
     * Get the value of a counter metric.
     * 
     * @param name Counter metric name
     * @return Current counter value or 0 if not found
     */
    private double getCounterValue(String name) {
        return meterRegistry.find(name).counter() != null ? 
               meterRegistry.find(name).counter().count() : 0.0;
    }
    
    /**
     * Get statistics for a timer metric.
     * 
     * @param name Timer metric name
     * @return Map of timer statistics
     */
    private Map<String, Object> getTimerStats(String name) {
        Map<String, Object> stats = new HashMap<>();
        
        io.micrometer.core.instrument.Timer timer = meterRegistry.find(name).timer();
        if (timer != null) {
            stats.put("count", timer.count());
            stats.put("totalTime", timer.totalTime(java.util.concurrent.TimeUnit.SECONDS));
            stats.put("mean", timer.mean(java.util.concurrent.TimeUnit.MILLISECONDS));
            stats.put("max", timer.max(java.util.concurrent.TimeUnit.MILLISECONDS));
            
            // Get statistics instead of deprecated percentiles
            Map<String, Double> statistics = new HashMap<>();
            statistics.put("mean", timer.mean(java.util.concurrent.TimeUnit.MILLISECONDS));
            statistics.put("max", timer.max(java.util.concurrent.TimeUnit.MILLISECONDS));
            statistics.put("totalTime", timer.totalTime(java.util.concurrent.TimeUnit.MILLISECONDS));
            
            stats.put("statistics", statistics);
        }
        
        return stats;
    }
    
    /**
     * Get the value of a gauge metric.
     * 
     * @param name Gauge metric name
     * @return Current gauge value or 0 if not found
     */
    private double getGaugeValue(String name) {
        io.micrometer.core.instrument.Gauge gauge = meterRegistry.find(name).gauge();
        return gauge != null ? gauge.value() : 0.0;
    }
    
    /**
     * Get memory usage statistics.
     * 
     * @return Map of memory usage statistics
     */
    private Map<String, Object> getMemoryStats() {
        Map<String, Object> memory = new HashMap<>();
        
        Runtime runtime = Runtime.getRuntime();
        long maxMemory = runtime.maxMemory();
        long allocatedMemory = runtime.totalMemory();
        long freeMemory = runtime.freeMemory();
        long usedMemory = allocatedMemory - freeMemory;
        
        memory.put("total", maxMemory / (1024 * 1024) + " MB");
        memory.put("used", usedMemory / (1024 * 1024) + " MB");
        memory.put("free", freeMemory / (1024 * 1024) + " MB");
        memory.put("usagePercent", (double) usedMemory / maxMemory * 100);
        
        return memory;
    }
}
