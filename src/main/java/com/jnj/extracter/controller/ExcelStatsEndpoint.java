package com.jnj.extracter.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.actuate.endpoint.annotation.Endpoint;
import org.springframework.boot.actuate.endpoint.annotation.ReadOperation;
import org.springframework.stereotype.Component;
import io.micrometer.core.instrument.MeterRegistry;
import io.micrometer.core.instrument.Meter;
import io.micrometer.core.instrument.Timer;

import java.util.HashMap;
import java.util.Map;
import java.util.stream.StreamSupport;

/**
 * Custom actuator endpoint to expose Excel processing statistics.
 * Accessible via: /actuator/excel-stats
 */
@Component
@Endpoint(id = "excel-stats")
public class ExcelStatsEndpoint {

    private final MeterRegistry meterRegistry;
    
    @Autowired
    public ExcelStatsEndpoint(MeterRegistry meterRegistry) {
        this.meterRegistry = meterRegistry;
    }
    
    @ReadOperation
    public Map<String, Object> getExcelStats() {
        Map<String, Object> stats = new HashMap<>();
        
        // Get counters
        stats.put("filesProcessed", getCounterValue("excel.files.processed"));
        stats.put("rowsProcessed", getCounterValue("excel.rows.processed"));
        stats.put("errors", getCounterValue("excel.processing.errors"));
        
        // Get timer statistics
        Map<String, Object> timers = new HashMap<>();
        timers.put("fileProcessing", getTimerStats("excel.file.processing.time"));
        timers.put("sheetProcessing", getTimerStats("excel.sheet.processing.time"));
        stats.put("timers", timers);
        
        // Get file type statistics
        Map<String, Object> fileTypes = new HashMap<>();
        StreamSupport.stream(meterRegistry.getMeters().spliterator(), false)
            .filter(meter -> meter.getId().getName().equals("excel.files.processed.by.type"))
            .forEach(meter -> {
                String fileType = meter.getId().getTag("fileType");
                if (fileType != null) {
                    fileTypes.put(fileType, getCounterValue(meter));
                }
            });
        stats.put("fileTypes", fileTypes);
        
        // Get error statistics
        Map<String, Object> errorTypes = new HashMap<>();
        StreamSupport.stream(meterRegistry.getMeters().spliterator(), false)
            .filter(meter -> meter.getId().getName().equals("excel.processing.errors.by.type"))
            .forEach(meter -> {
                String errorType = meter.getId().getTag("errorType");
                if (errorType != null) {
                    errorTypes.put(errorType, getCounterValue(meter));
                }
            });
        stats.put("errorTypes", errorTypes);
        
        // Get system stats
        Map<String, Object> system = new HashMap<>();
        Runtime runtime = Runtime.getRuntime();
        system.put("cpus", runtime.availableProcessors());
        system.put("memoryTotal", runtime.totalMemory() / (1024 * 1024));
        system.put("memoryFree", runtime.freeMemory() / (1024 * 1024));
        system.put("memoryMax", runtime.maxMemory() / (1024 * 1024));
        stats.put("system", system);
        
        return stats;
    }
    
    private double getCounterValue(String name) {
        try {
            return meterRegistry.find(name).counter() != null 
                ? meterRegistry.find(name).counter().count() 
                : 0.0;
        } catch (Exception e) {
            return 0.0;
        }
    }
    
    private double getCounterValue(Meter meter) {
        try {
            return meterRegistry.find(meter.getId().getName())
                .tags(meter.getId().getTagsAsIterable())
                .counter() != null
                    ? meterRegistry.find(meter.getId().getName())
                        .tags(meter.getId().getTagsAsIterable())
                        .counter()
                        .count()
                    : 0.0;
        } catch (Exception e) {
            return 0.0;
        }
    }
    
    private Map<String, Object> getTimerStats(String name) {
        Map<String, Object> stats = new HashMap<>();
        
        try {
            Timer timer = meterRegistry.find(name).timer();
            
            if (timer != null) {
                stats.put("count", timer.count());
                stats.put("totalTimeSeconds", timer.totalTime(java.util.concurrent.TimeUnit.SECONDS));
                stats.put("maxTimeSeconds", timer.max(java.util.concurrent.TimeUnit.SECONDS));
                stats.put("meanTimeSeconds", timer.mean(java.util.concurrent.TimeUnit.SECONDS));
            } else {
                stats.put("count", 0);
                stats.put("totalTimeSeconds", 0.0);
                stats.put("maxTimeSeconds", 0.0);
                stats.put("meanTimeSeconds", 0.0);
            }
        } catch (Exception e) {
            stats.put("count", 0);
            stats.put("totalTimeSeconds", 0.0);
            stats.put("maxTimeSeconds", 0.0);
            stats.put("meanTimeSeconds", 0.0);
            stats.put("error", "Failed to retrieve timer stats: " + e.getMessage());
        }
        
        return stats;
    }
}
