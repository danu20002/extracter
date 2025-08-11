package com.jnj.extracter.api.exception;

/**
 * Exception thrown when input validation fails.
 */
public class ValidationException extends BaseException {
    
    public ValidationException(String message) {
        super(message, "VALIDATION_ERROR");
    }
    
    public ValidationException(String message, Throwable cause) {
        super(message, "VALIDATION_ERROR", cause);
    }
    
    public ValidationException(String field, String message) {
        super(String.format("Invalid value for %s: %s", field, message), "VALIDATION_ERROR");
    }
}
