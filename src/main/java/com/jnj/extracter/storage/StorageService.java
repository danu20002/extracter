package com.jnj.extracter.storage;

import java.io.File;
import java.io.InputStream;

public interface StorageService {
    /**
     * Save an attachment to storage
     * @param fileName The name of the file
     * @param inputStream The file data
     * @return The saved file or a reference string (e.g., DMS id)
     */
    Object save(String fileName, InputStream inputStream) throws Exception;
}
