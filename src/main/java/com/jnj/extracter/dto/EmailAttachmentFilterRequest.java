package com.jnj.extracter.dto;

import java.time.LocalDate;

public class EmailAttachmentFilterRequest {
    private String sender;
    private String subjectContains;
    private LocalDate startDate;
    private LocalDate endDate;
    private String saveDir;
    // Add more fields as needed

    public String getSender() { return sender; }
    public void setSender(String sender) { this.sender = sender; }

    public String getSubjectContains() { return subjectContains; }
    public void setSubjectContains(String subjectContains) { this.subjectContains = subjectContains; }

    public LocalDate getStartDate() { return startDate; }
    public void setStartDate(LocalDate startDate) { this.startDate = startDate; }

    public LocalDate getEndDate() { return endDate; }
    public void setEndDate(LocalDate endDate) { this.endDate = endDate; }

    public String getSaveDir() { return saveDir; }
    public void setSaveDir(String saveDir) { this.saveDir = saveDir; }
}
