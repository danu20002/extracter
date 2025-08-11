package com.jnj.extracter.api.exception;

/**
 * Base exception class for all application exceptions.
 * All custom exceptions should extend this class.
 */
public class BaseException extends RuntimeException {
    
    private final String code;
    
    public BaseException(String message) {
        super(message);
        this.code = "GENERAL_ERROR";
    }
    
    public BaseException(String message, Throwable cause) {
        super(message, cause);
        this.code = "GENERAL_ERROR";
    }
    
    public BaseException(String message, String code) {
        super(message);
        this.code = code;
    }
    
    public BaseException(String message, String code, Throwable cause) {
        super(message, cause);
        this.code = code;
    }
    
    public String getCode() {
        return code;
    }
}
