package com.jnj.extracter.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

/**
 * Configuration properties for file storage.
 */
@Configuration
@ConfigurationProperties(prefix = "application.storage")
@Data
public class StorageProperties {
    /**
     * Root directory for file storage.
     */
    private String location = "excel";
    
    /**
     * Maximum age of files in days before cleanup.
     */
    private int maxAgeDays = 30;
    
    /**
     * Maximum file size in bytes.
     */
    private long maxFileSize = 10485760; // Default: 10 MB
}
