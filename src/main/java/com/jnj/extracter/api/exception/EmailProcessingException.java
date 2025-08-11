package com.jnj.extracter.api.exception;

/**
 * Exception thrown when there's an error processing emails or attachments.
 */
public class EmailProcessingException extends BaseException {
    
    public EmailProcessingException(String message) {
        super(message);
    }
    
    public EmailProcessingException(String message, Throwable cause) {
        super(message, cause);
    }
}
