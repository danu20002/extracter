package com.jnj.extracter.service.impl;

import com.jnj.extracter.domain.model.ExcelFileInfo;
import com.jnj.extracter.service.contract.EmailReaderService;
import com.jnj.extracter.service.contract.StorageService;
import jakarta.mail.*;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Implementation of EmailReaderService for reading email attachments.
 */
@Service
@RequiredArgsConstructor
@Slf4j
public class EmailReaderServiceImpl implements EmailReaderService {

    private final StorageService storageService;

    @Value("${attachment.save.dir:excel/temp}")
    private String defaultSaveDir;

    @Override
    public List<Object> fetchAttachmentsFromEmail(String host, String port, String user, String password, String saveDir) throws Exception {
        log.info("Fetching email attachments from {}:{} for user {}", host, port, user);
        
        if (saveDir == null || saveDir.isEmpty()) {
            saveDir = defaultSaveDir;
        }
        
        List<Object> savedAttachments = new ArrayList<>();
        
        // Set up mail properties
        Properties props = new Properties();
        props.put("mail.store.protocol", "imaps");
        props.put("mail.imaps.host", host);
        props.put("mail.imaps.port", port);
        props.put("mail.imaps.ssl.enable", "true");
        
        // Get session
        Session session = Session.getInstance(props, null);
        
        try (Store store = session.getStore()) {
            store.connect(host, user, password);
            
            // Get inbox
            Folder inbox = store.getFolder("INBOX");
            inbox.open(Folder.READ_ONLY);
            
            log.info("Connected to inbox with {} messages", inbox.getMessageCount());
            
            // Get messages
            Message[] messages = inbox.getMessages();
            
            for (Message message : messages) {
                log.debug("Processing email: {}", message.getSubject());
                
                // Check if message has attachments
                if (message.getContentType().contains("multipart")) {
                    Multipart multipart = (Multipart) message.getContent();
                    
                    for (int i = 0; i < multipart.getCount(); i++) {
                        BodyPart bodyPart = multipart.getBodyPart(i);
                        
                        if (Part.ATTACHMENT.equalsIgnoreCase(bodyPart.getDisposition())) {
                            String fileName = bodyPart.getFileName();
                            
                            // Only process Excel files
                            if (fileName.toLowerCase().endsWith(".xlsx") || 
                                fileName.toLowerCase().endsWith(".xls") ||
                                fileName.toLowerCase().endsWith(".xlsb")) {
                                
                                log.info("Found Excel attachment: {}", fileName);
                                
                                // Create a temporary file and save the attachment
                                File tempFile = File.createTempFile("email-attachment-", fileName);
                                try (InputStream inputStream = bodyPart.getInputStream();
                                     java.io.FileOutputStream outputStream = new java.io.FileOutputStream(tempFile)) {
                                    byte[] buffer = new byte[8192];
                                    int bytesRead;
                                    while ((bytesRead = inputStream.read(buffer)) != -1) {
                                        outputStream.write(buffer, 0, bytesRead);
                                    }
                                }
                                
                                // Store the file using the StorageService
                                String savedFileName = storageService.store(tempFile);
                                savedAttachments.add(savedFileName);
                                
                                // Delete the temporary file
                                tempFile.delete();
                            }
                        }
                    }
                }
            }
            
            inbox.close(false);
        }
        
        log.info("Fetched {} attachments", savedAttachments.size());
        return savedAttachments;
    }

    @Override
    public List<ExcelFileInfo> getAttachments(
            String subject,
            String sender,
            LocalDate fromDate,
            LocalDate toDate,
            boolean includeSubfolders) {
        
        log.info("Getting attachments with filter - subject: {}, sender: {}", subject, sender);
        
        // In a real implementation, this would query emails based on the filters
        // For now, we'll return files from the storage directory
        
        try {
            File directory = new File(defaultSaveDir);
            
            if (!directory.exists() || !directory.isDirectory()) {
                log.warn("Storage directory does not exist: {}", defaultSaveDir);
                return new ArrayList<>();
            }
            
            List<File> files = Arrays.asList(directory.listFiles((dir, name) -> {
                return name.toLowerCase().endsWith(".xlsx") || 
                       name.toLowerCase().endsWith(".xls") ||
                       name.toLowerCase().endsWith(".xlsb");
            }));
            
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            
            return files.stream()
                .map(file -> new ExcelFileInfo(
                    file.getName(),
                    file.length(),
                    file.getAbsolutePath(),
                    sdf.format(new Date(file.lastModified()))
                ))
                .collect(Collectors.toList());
                
        } catch (Exception e) {
            log.error("Error getting attachments", e);
            return new ArrayList<>();
        }
    }
}
