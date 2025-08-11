package com.jnj.extracter.service.contract;

/**
 * Service interface for cleaning up old files.
 */
public interface FileSystemCleanupService {
    
    /**
     * Cleans up old files based on configured retention policy.
     */
    void cleanupOldFiles();
}
