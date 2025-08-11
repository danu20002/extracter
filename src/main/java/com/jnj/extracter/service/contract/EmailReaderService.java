package com.jnj.extracter.service.contract;

import com.jnj.extracter.domain.model.ExcelFileInfo;

import java.time.LocalDate;
import java.util.List;

/**
 * Service interface for reading email attachments.
 */
public interface EmailReaderService {
	/**
	 * Fetch attachments from email and save to directory or DMS.
	 * 
	 * @param host IMAP host
	 * @param port IMAP port
	 * @param user Email username
	 * @param password Email password
	 * @param saveDir Directory to save attachments (if local)
	 * @return List of saved references (File or DMS id)
	 * @throws Exception if an error occurs during fetching or saving
	 */
	List<Object> fetchAttachmentsFromEmail(
		String host, 
		String port, 
		String user, 
		String password, 
		String saveDir
	) throws Exception;
	
	/**
	 * Get email attachments based on filter criteria.
	 * 
	 * @param subject Email subject filter
	 * @param sender Sender email address filter
	 * @param fromDate Start date for emails
	 * @param toDate End date for emails
	 * @param includeSubfolders Whether to include subfolders in search
	 * @return List of Excel file information
	 */
	List<ExcelFileInfo> getAttachments(
		String subject,
		String sender,
		LocalDate fromDate,
		LocalDate toDate,
		boolean includeSubfolders
	);
}
