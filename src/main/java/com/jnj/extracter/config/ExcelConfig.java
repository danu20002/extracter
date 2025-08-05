package com.jnj.extracter.config;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.context.annotation.Configuration;

import jakarta.annotation.PostConstruct;


@Configuration
public class ExcelConfig {

    @PostConstruct
    public void init() {
        // Configure Apache POI settings to handle problematic Excel files
        
        // Set minimum inflate ratio to allow for highly compressed files (prevents zip bomb detection)
        // Use a more permissive ratio to handle extreme compression
        ZipSecureFile.setMinInflateRatio(0.0001);
        
        // Allow larger entries (100MB) - increased from 50MB to handle larger worksheets
        ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024);
        
        // Increase maximum number of records when handling SAX events
        // Useful for very large Excel files
        System.setProperty("org.apache.poi.xssf.max_rows", "1000000");
        
        System.out.println("Excel configuration initialized with relaxed security settings");
    }
}
