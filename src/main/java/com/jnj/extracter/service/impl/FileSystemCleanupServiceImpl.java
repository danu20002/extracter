package com.jnj.extracter.service.impl;

import com.jnj.extracter.config.StorageProperties;
import com.jnj.extracter.service.contract.FileSystemCleanupService;
import com.jnj.extracter.service.contract.StorageService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.stream.Stream;

/**
 * Implementation of FileSystemCleanupService that cleans up old files.
 */
@Service
@Slf4j
public class FileSystemCleanupServiceImpl implements FileSystemCleanupService {

    private final StorageService storageService;
    private final StorageProperties storageProperties;
    
    public FileSystemCleanupServiceImpl(StorageService storageService, StorageProperties storageProperties) {
        this.storageService = storageService;
        this.storageProperties = storageProperties;
        log.info("FileSystemCleanupServiceImpl initialized");
    }
    
    @Override
    @Scheduled(cron = "${application.cleanup.cron:0 0 3 * * ?}") // Default: 3 AM daily
    public void cleanupOldFiles() {
        log.info("Starting cleanup of old files");
        
        int maxAgeDays = storageProperties.getMaxAgeDays();
        if (maxAgeDays <= 0) {
            log.info("File cleanup disabled (maxAgeDays <= 0)");
            return;
        }
        
        Path rootLocation = Paths.get(storageProperties.getLocation());
        Instant cutoffTime = Instant.now().minus(maxAgeDays, ChronoUnit.DAYS);
        
        try (Stream<Path> pathStream = Files.walk(rootLocation)) {
            long deletedCount = pathStream
                .filter(Files::isRegularFile)
                .filter(path -> {
                    try {
                        BasicFileAttributes attr = Files.readAttributes(path, BasicFileAttributes.class);
                        return attr.creationTime().toInstant().isBefore(cutoffTime);
                    } catch (IOException e) {
                        log.warn("Could not read attributes of file: {}", path, e);
                        return false;
                    }
                })
                .map(path -> {
                    try {
                        String filename = path.getFileName().toString();
                        log.debug("Deleting old file: {}", filename);
                        Files.delete(path);
                        return true;
                    } catch (IOException e) {
                        log.warn("Failed to delete file: {}", path, e);
                        return false;
                    }
                })
                .filter(deleted -> deleted)
                .count();
                
            log.info("Cleanup complete. Deleted {} files older than {} days", deletedCount, maxAgeDays);
        } catch (IOException e) {
            log.error("Error during file cleanup", e);
        }
    }
}
