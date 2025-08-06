package com.jnj.extracter.config;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;

import lombok.Getter;

@Configuration
@PropertySource("classpath:application.properties")
@Getter
public class ExcelProcessingConfig {
    
    @Value("${excel.folder.path:excel}")
    private String excelFolderPath;
    
    @Value("${excel.max.file.size:100MB}")
    private String maxFileSize;
    
    @Value("${excel.processing.batch-size:1000}")
    private int batchSize;
    
    @Value("${excel.buffer.size:8192}")
    private int bufferSize;
    
    @Value("${excel.use.memory-mapped:true}")
    private boolean useMemoryMapped;
    
    @Value("${excel.parallel.processing:true}")
    private boolean parallelProcessing;
    
    @Value("${excel.thread.pool.size:4}")
    private int threadPoolSize;
}
