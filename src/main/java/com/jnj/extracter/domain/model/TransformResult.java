package com.jnj.extracter.domain.model;

import lombok.Data;

/**
 * Represents the result of a file transformation operation.
 */
@Data
public class TransformResult {
    private String sourceFilename;
    private String transformedFilename;
    private String transformType;
    private long processingTimeMs;
}
