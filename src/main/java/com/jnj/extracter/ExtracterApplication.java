package com.jnj.extracter;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.scheduling.annotation.EnableAsync;
import org.springframework.scheduling.concurrent.ThreadPoolTaskExecutor;

import java.util.concurrent.Executor;

@SpringBootApplication
@EnableAsync
public class ExtracterApplication {

    public static void main(String[] args) {
        // Set global POI security settings
        ZipSecureFile.setMinInflateRatio(0.0001);
        ZipSecureFile.setMaxEntrySize(100 * 1024 * 1024); // 100MB
        System.setProperty("org.apache.poi.xssf.parsemode", "tolerant");
        System.setProperty("org.apache.poi.ooxml.strict", "false");
        
        // For large heap allocations
        System.setProperty("sun.java2d.cmm", "sun.java2d.cmm.kcms.KcmsServiceProvider");
        
        // Run the application
        SpringApplication.run(ExtracterApplication.class, args);
    }
    
    /**
     * Configure an async executor for background tasks
     */
    @Bean
    public Executor taskExecutor() {
        ThreadPoolTaskExecutor executor = new ThreadPoolTaskExecutor();
        executor.setCorePoolSize(4);
        executor.setMaxPoolSize(10);
        executor.setQueueCapacity(100);
        executor.setThreadNamePrefix("ExcelProcessor-");
        executor.setWaitForTasksToCompleteOnShutdown(true);
        executor.initialize();
        return executor;
    }
}
