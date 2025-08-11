package com.jnj.extracter.api.exception;

/**
 * Exception thrown when file storage operations fail.
 */
public class FileStorageException extends BaseException {
    
    public FileStorageException(String message) {
        super(message, "FILE_STORAGE_ERROR");
    }
    
    public FileStorageException(String message, Throwable cause) {
        super(message, "FILE_STORAGE_ERROR", cause);
    }
    
    public FileStorageException(String message, String code) {
        super(message, code);
    }
    
    public FileStorageException(String message, String code, Throwable cause) {
        super(message, code, cause);
    }
}
