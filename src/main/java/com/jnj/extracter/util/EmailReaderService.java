package com.jnj.extracter.util;

import java.util.List;

public interface EmailReaderService {
	/**
	 * Fetch attachments from email and save to directory or DMS
	 * @param host IMAP host
	 * @param port IMAP port
	 * @param user Email username
	 * @param password Email password
	 * @param saveDir Directory to save attachments (if local)
	 * @return List of saved references (File or DMS id)
	 */
	List<Object> fetchAttachmentsFromEmail(String host, String port, String user, String password, String saveDir) throws Exception;
}
