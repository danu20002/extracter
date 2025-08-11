
package com.jnj.extracter.controller;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.jnj.extracter.dto.EmailAttachmentFilterRequest;
import com.jnj.extracter.serviceImpl.EmailReaderImpl;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;

import java.io.File;
import java.util.*;
import java.util.stream.Collectors;

@RestController
@RequestMapping("/api/email-attachments")
public class EmailAttachmentController {
    private static final Logger logger = LoggerFactory.getLogger(EmailAttachmentController.class);
    @Autowired
    private EmailReaderImpl emailReader;
    @Value("${email.imap.host}")
    private String host;
    @Value("${email.imap.port}")
    private String port;
    @Value("${email.imap.user}")
    private String user;
    @Value("${email.imap.password}")
    private String password;

    // POST: Fetch attachments with filters
    @PostMapping("/fetch")
    public ResponseEntity<List<String>> fetchAttachments(@RequestBody EmailAttachmentFilterRequest filter) {
        String saveDir = filter.getSaveDir() != null ? filter.getSaveDir() : "excel/temp";
        // Log all email and filter details (do not log password in production!)
        logger.info("Fetching email attachments with settings: host={}, port={}, user={}, saveDir={}, filter={}", host, port, user, saveDir, filter);
        try {
            List<Object> refs = emailReader.fetchAttachmentsFromEmail(host, port, user, password, saveDir);
            // Convert all references to string for response (file name or DMS id)
            List<String> refStrings = refs.stream().map(ref -> {
                if (ref instanceof File) return ((File)ref).getName();
                return ref.toString();
            }).collect(Collectors.toList());
            return ResponseEntity.ok(refStrings);
        } catch (Exception e) {
            logger.error("Error fetching email attachments", e);
            return ResponseEntity.badRequest().body(Collections.singletonList(e.getMessage()));
        }
    }

    // GET: List all downloaded attachments
    @GetMapping
    public ResponseEntity<List<String>> listAttachments(@RequestParam(defaultValue = "excel/temp") String dir) {
        File folder = new File(dir);
        if (!folder.exists() || !folder.isDirectory()) return ResponseEntity.ok(Collections.emptyList());
        String[] files = folder.list();
        return ResponseEntity.ok(files != null ? Arrays.asList(files) : Collections.emptyList());
    }

    // GET: Get metadata for a single attachment
    @GetMapping("/{fileName}")
    public ResponseEntity<Map<String, Object>> getAttachment(@PathVariable String fileName, @RequestParam(defaultValue = "excel/temp") String dir) {
        File file = new File(dir, fileName);
        if (!file.exists()) return ResponseEntity.notFound().build();
        Map<String, Object> meta = new HashMap<>();
        meta.put("name", file.getName());
        meta.put("size", file.length());
        meta.put("lastModified", file.lastModified());
        return ResponseEntity.ok(meta);
    }

    // DELETE: Delete an attachment
    @DeleteMapping("/{fileName}")
    public ResponseEntity<?> deleteAttachment(@PathVariable String fileName, @RequestParam(defaultValue = "excel/temp") String dir) {
        File file = new File(dir, fileName);
        if (file.exists() && file.delete()) {
            return ResponseEntity.ok().build();
        } else {
            return ResponseEntity.notFound().build();
        }
    }
}
