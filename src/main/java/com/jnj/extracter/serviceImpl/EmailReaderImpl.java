package com.jnj.extracter.serviceImpl;


import com.jnj.extracter.util.EmailReaderService;
import jakarta.mail.Session;
import jakarta.mail.Store;
import jakarta.mail.Folder;
import jakarta.mail.Message;
import jakarta.mail.Multipart;
import jakarta.mail.BodyPart;
import jakarta.mail.Part;
import com.jnj.extracter.storage.StorageService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import java.io.File;
import java.util.*;

import org.springframework.stereotype.Service;

@Service
public class EmailReaderImpl implements EmailReaderService {

	@Autowired
	private StorageService storageService;

	@Value("${attachment.save.dir:excel/temp}")
	private String defaultSaveDir;

	@Override
	public List<Object> fetchAttachmentsFromEmail(String host, String port, String user, String password, String saveDir) throws Exception {
		List<Object> savedRefs = new ArrayList<>();
		Properties properties = new Properties();
		properties.put("mail.store.protocol", "imaps");
		properties.put("mail.imaps.host", host);
		properties.put("mail.imaps.port", port);

		Session session = Session.getDefaultInstance(properties);
		Store store = session.getStore("imaps");
		store.connect(host, user, password);

		Folder inbox = store.getFolder("INBOX");
		inbox.open(Folder.READ_ONLY);

		Message[] messages = inbox.getMessages();

		for (Message message : messages) {
			if (message.isMimeType("multipart/*")) {
				Multipart multipart = (Multipart) message.getContent();
				for (int i = 0; i < multipart.getCount(); i++) {
					BodyPart part = multipart.getBodyPart(i);
					if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
						String fileName = part.getFileName();
						try (java.io.InputStream is = part.getInputStream()) {
							Object ref = storageService.save(fileName, is);
							savedRefs.add(ref);
						}
					}
				}
			}
		}

		inbox.close(false);
		store.close();
		return savedRefs;
	}
}
