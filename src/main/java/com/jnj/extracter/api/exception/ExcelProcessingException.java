package com.jnj.extracter.api.exception;

/**
 * Exception thrown when Excel processing fails.
 */
public class ExcelProcessingException extends BaseException {
    
    public ExcelProcessingException(String message) {
        super(message, "EXCEL_PROCESSING_ERROR");
    }
    
    public ExcelProcessingException(String message, Throwable cause) {
        super(message, "EXCEL_PROCESSING_ERROR", cause);
    }
    
    public ExcelProcessingException(String message, String code) {
        super(message, code);
    }
    
    public ExcelProcessingException(String message, String code, Throwable cause) {
        super(message, code, cause);
    }
}
