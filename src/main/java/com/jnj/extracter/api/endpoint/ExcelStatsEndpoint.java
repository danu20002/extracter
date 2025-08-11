package com.jnj.extracter.api.endpoint;

import io.micrometer.core.instrument.MeterRegistry;
import org.springframework.boot.actuate.endpoint.annotation.Endpoint;
import org.springframework.boot.actuate.endpoint.annotation.ReadOperation;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.Map;

/**
 * Custom actuator endpoint to expose Excel processing statistics.
 * Accessible via: /actuator/excel-stats
 * 
 * This class provides metrics through the Spring Boot Actuator.
 */
@Component
@Endpoint(id = "excel-stats")
public class ExcelStatsEndpoint {

    private final MeterRegistry meterRegistry;
    
    public ExcelStatsEndpoint(MeterRegistry meterRegistry) {
        this.meterRegistry = meterRegistry;
    }
    
    /**
     * Read operation for the actuator endpoint.
     * 
     * @return Map containing various Excel processing metrics
     */
    @ReadOperation
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
}
