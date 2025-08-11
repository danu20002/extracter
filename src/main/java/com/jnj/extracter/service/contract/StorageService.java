package com.jnj.extracter.service.contract;

import org.springframework.core.io.Resource;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

/**
 * Service interface for file storage operations.
 */
public interface StorageService {
    
    /**
     * Initialize the storage.
     */
    void init();
    
    /**
     * Store a file from a MultipartFile.
     * 
     * @param file The file to store
     * @return The stored filename
     */
    String store(MultipartFile file);
    
    /**
     * Store a file from a File object.
     * 
     * @param file The file to store
     * @return The stored filename
     * @throws IOException If there's an error storing the file
     */
    String store(File file) throws IOException;
    
    /**
     * Load all files from storage.
     * 
     * @return List of file paths
     * @throws IOException If there's an error reading the files
     */
    List<Path> loadAll() throws IOException;
    
    /**
     * Load a file by name.
     * 
     * @param filename The name of the file
     * @return The file path
     */
    Path load(String filename);
    
    /**
     * Load a file as a Resource.
     * 
     * @param filename The name of the file
     * @return The file resource
     */
    Resource loadAsResource(String filename);
    
    /**
     * Delete a file.
     * 
     * @param filename The name of the file to delete
     */
    void delete(String filename);
    
    /**
     * Delete all files.
     */
    void deleteAll();
    
    /**
     * Get the size of a file.
     * 
     * @param path The file path
     * @return The file size in bytes
     * @throws IOException If there's an error reading the file
     */
    long getFileSize(Path path) throws IOException;
    
    /**
     * Get the last modified time of a file.
     * 
     * @param path The file path
     * @return The last modified time
     * @throws IOException If there's an error reading the file
     */
    LocalDateTime getLastModifiedTime(Path path) throws IOException;
    
    /**
     * Get the root location of the storage.
     * 
     * @return The root location path
     */
    Path getRootLocation();
}
