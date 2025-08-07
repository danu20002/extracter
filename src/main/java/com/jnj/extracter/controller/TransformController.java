package com.jnj.extracter.controller;

import com.jnj.extracter.transform.Transform;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/transform")
public class TransformController {

    @Autowired
    private Transform transformService;

    @GetMapping("/journal")
    public String transformJournal() {
        return transformService.transformJournalWithMaster();
    }

    @GetMapping("/generate-journal")
    public ResponseEntity<ByteArrayResource> generateJournalFromMaster() {
        byte[] data = transformService.generateJournalFromMaster();
        if (data == null || data.length == 0) {
            return ResponseEntity.notFound().build();
        }
        ByteArrayResource resource = new ByteArrayResource(data);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Journal_transformed.xlsx")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .contentLength(data.length)
                .body(resource);
    }
}
