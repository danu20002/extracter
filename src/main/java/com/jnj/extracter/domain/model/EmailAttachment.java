package com.jnj.extracter.domain.model;

import lombok.Data;

import java.util.Date;

/**
 * Represents an email attachment, typically an Excel file.
 */
@Data
public class EmailAttachment {
    private String originalFileName;
    private String storedFileName;
    private String subject;
    private String from;
    private Date sentDate;
    private long size;
}
