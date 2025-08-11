package com.jnj.extracter.api.dto;

import jakarta.validation.constraints.Email;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Past;
import jakarta.validation.constraints.Pattern;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;

/**
 * Request DTO for filtering email attachments.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class EmailAttachmentFilterRequest {
    
    /**
     * Email sender address to filter by
     */
    @Email(message = "Invalid email format")
    private String sender;
    
    /**
     * Text that should be contained in the email subject
     */
    @NotBlank(message = "Subject filter cannot be empty")
    private String subjectContains;
    
    /**
     * Start date for the date range filter
     */
    @Past(message = "Start date must be in the past")
    private LocalDate startDate;
    
    /**
     * End date for the date range filter
     */
    private LocalDate endDate;
    
    /**
     * Directory to save attachments
     */
    @Pattern(regexp = "^[^<>:\"\\\\|?*]*$", message = "Directory path contains invalid characters")
    private String saveDir;
    
    /**
     * Whether to include subfolders in the search
     */
    @Builder.Default
    private boolean includeSubfolders = false;
}
