package com.jnj.extracter.web.controller;

import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.ExcelFileInfo;
import com.jnj.extracter.domain.model.ExcelProcessingResult;
import com.jnj.extracter.service.contract.ExcelService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Web controller for Excel data extraction with Thymeleaf views.
 * This controller handles the web UI interactions for the Excel processing application.
 */
@Controller
@RequestMapping("/excel")
@RequiredArgsConstructor
@Slf4j
public class ExcelWebController {

    private final ExcelService excelService;

    /**
     * Main dashboard page.
     * 
     * @param model The model for the view
     * @return The dashboard view name
     */
    @GetMapping("/dashboard")
    public String dashboard(Model model) {
        model.addAttribute("title", "Excel Extractor Dashboard");
        return "excel/dashboard";
    }
    
    /**
     * List all Excel files in the system.
     * 
     * @param model The model for the view
     * @return The files view name
     */
    @GetMapping("/files")
    public String listFiles(Model model) {
        List<File> files = excelService.getExcelFiles();
        
        List<ExcelFileInfo> fileInfos = files.stream()
            .map(file -> {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                String lastModified = sdf.format(new Date(file.lastModified()));
                
                return new ExcelFileInfo(
                    file.getName(),
                    file.length(),
                    file.getAbsolutePath(),
                    lastModified
                );
            })
            .collect(Collectors.toList());
        
        model.addAttribute("files", fileInfos);
        model.addAttribute("title", "Excel Files");
        
        log.info("Listing {} Excel files", fileInfos.size());
        
        return "excel/files";
    }
    
    /**
     * Handle file upload.
     * 
     * @param file The uploaded file
     * @param redirectAttributes Attributes for the redirect
     * @return Redirect to files page
     */
    @PostMapping("/upload")
    public String uploadFile(
            @RequestParam("file") MultipartFile file,
            RedirectAttributes redirectAttributes) {
        
        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("error", "Please select a file to upload");
            return "redirect:/excel/files";
        }
        
        try {
            // Save the file
            String fileName = file.getOriginalFilename();
            File savedFile = new File("excel/" + fileName);
            
            // Ensure directory exists
            savedFile.getParentFile().mkdirs();
            
            try (FileOutputStream fos = new FileOutputStream(savedFile)) {
                fos.write(file.getBytes());
            }
            
            log.info("Successfully uploaded file: {}", fileName);
            
            redirectAttributes.addFlashAttribute("success", 
                    "File uploaded successfully: " + fileName);
            
        } catch (IOException e) {
            log.error("Failed to upload file", e);
            redirectAttributes.addFlashAttribute("error", 
                    "Failed to upload file: " + e.getMessage());
        }
        
        return "redirect:/excel/files";
    }
    
    /**
     * View a specific Excel file.
     * 
     * @param fileName Name of the file
     * @param model The model for the view
     * @return The view-file view name
     */
    @GetMapping("/file/{fileName}")
    public String viewFile(@PathVariable String fileName, Model model) {
        try {
            List<File> files = excelService.getExcelFiles();
            
            File targetFile = files.stream()
                .filter(f -> f.getName().equals(fileName))
                .findFirst()
                .orElse(null);
            
            if (targetFile == null) {
                model.addAttribute("error", "File not found: " + fileName);
                return "excel/error";
            }
            
            List<String> sheetNames = excelService.getSheetNames(targetFile);
            
            // Create file info
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            String lastModified = sdf.format(new Date(targetFile.lastModified()));
            
            ExcelFileInfo fileInfo = new ExcelFileInfo(
                targetFile.getName(),
                targetFile.length(),
                targetFile.getAbsolutePath(),
                lastModified
            );
            
            model.addAttribute("file", fileInfo);
            model.addAttribute("sheetNames", sheetNames);
            model.addAttribute("title", "File: " + fileName);
            
            log.info("Viewing file {} with {} sheets", fileName, sheetNames.size());
            
            return "excel/view-file";
            
        } catch (Exception e) {
            log.error("Error viewing file: {}", fileName, e);
            model.addAttribute("error", "Error viewing file: " + e.getMessage());
            return "excel/error";
        }
    }
}
