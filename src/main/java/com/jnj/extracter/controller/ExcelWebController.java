package com.jnj.extracter.controller;

import com.jnj.extracter.entity.ExcelData;
import com.jnj.extracter.entity.ExcelFileInfo;
import com.jnj.extracter.entity.ExcelProcessingResult;
import com.jnj.extracter.service.ExcelService;
import lombok.RequiredArgsConstructor;
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
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Web controller for Excel data extraction with Thymeleaf views
 */
@Controller
@RequestMapping("/excel")
@RequiredArgsConstructor
public class ExcelWebController {

    private final ExcelService excelService;

    /**
     * Main dashboard page
     */
    @GetMapping
    public String dashboard(Model model) {
        List<File> files = excelService.getExcelFiles();
        List<ExcelFileInfo> fileInfoList = convertToFileInfoList(files);
        model.addAttribute("files", fileInfoList);
        return "excel/dashboard";
    }
    
    /**
     * Convert File objects to ExcelFileInfo objects to avoid Thymeleaf security restrictions
     */
    private List<ExcelFileInfo> convertToFileInfoList(List<File> files) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        return files.stream().map(file -> {
            ExcelFileInfo info = new ExcelFileInfo();
            info.setName(file.getName());
            info.setSize(file.length());
            info.setPath(file.getAbsolutePath());
            info.setLastModified(dateFormat.format(new Date(file.lastModified())));
            return info;
        }).collect(Collectors.toList());
    }

    /**
     * Handle file upload
     */
    @PostMapping("/upload")
    public String handleFileUpload(@RequestParam("file") MultipartFile file,
                                   RedirectAttributes redirectAttributes) {
        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("error", "Please select a file to upload");
            return "redirect:/excel";
        }

        try {
            // Get the Excel directory
            List<File> existingFiles = excelService.getExcelFiles();
            String uploadDir = "";
            if (!existingFiles.isEmpty()) {
                uploadDir = existingFiles.get(0).getParentFile().getAbsolutePath();
            } else {
                // Default upload directory if no existing files
                uploadDir = Paths.get(System.getProperty("user.dir"), "excel").toString();
                new File(uploadDir).mkdirs();
            }

            // Save the file
            File destination = new File(uploadDir, file.getOriginalFilename());
            try (FileOutputStream os = new FileOutputStream(destination)) {
                os.write(file.getBytes());
            }

            redirectAttributes.addFlashAttribute("success", 
                "File uploaded successfully: " + file.getOriginalFilename());
            
        } catch (IOException e) {
            redirectAttributes.addFlashAttribute("error", 
                "Failed to upload file: " + e.getMessage());
        }

        return "redirect:/excel";
    }

    /**
     * View all Excel files
     */
    @GetMapping("/files")
    public String viewFiles(Model model) {
        List<File> files = excelService.getExcelFiles();
        List<ExcelFileInfo> fileInfoList = convertToFileInfoList(files);
        model.addAttribute("files", fileInfoList);
        return "excel/files";
    }

    /**
     * Extract and view data from a specific file
     */
    @GetMapping("/view/{fileName}")
    public String viewFileData(@PathVariable String fileName, Model model) {
        List<File> files = excelService.getExcelFiles();
        File targetFile = files.stream()
                .filter(file -> file.getName().equals(fileName))
                .findFirst()
                .orElse(null);

        if (targetFile == null) {
            model.addAttribute("error", "File not found: " + fileName);
            return "excel/error";
        }
        
        ExcelFileInfo fileInfo = new ExcelFileInfo();
        fileInfo.setName(targetFile.getName());
        fileInfo.setSize(targetFile.length());
        fileInfo.setPath(targetFile.getAbsolutePath());
        fileInfo.setLastModified(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date(targetFile.lastModified())));

        ExcelProcessingResult result = excelService.extractExcelFile(targetFile);
        model.addAttribute("file", fileInfo);
        model.addAttribute("result", result);
        return "excel/view-file";
    }

    /**
     * View data from a specific sheet
     */
    @GetMapping("/view/{fileName}/{sheetName}")
    public String viewSheetData(@PathVariable String fileName,
                               @PathVariable String sheetName,
                               Model model) {
        
        List<File> files = excelService.getExcelFiles();
        File targetFile = files.stream()
                .filter(file -> file.getName().equals(fileName))
                .findFirst()
                .orElse(null);

        if (targetFile == null) {
            model.addAttribute("error", "File not found: " + fileName);
            return "excel/error";
        }

        List<ExcelData> data = excelService.extractSheetData(targetFile, sheetName);
        
        // Extract column headers from the first row
        List<String> headers = new ArrayList<>();
        if (!data.isEmpty() && data.get(0).getData() != null) {
            headers.addAll(data.get(0).getData().keySet());
        }
        
        model.addAttribute("fileName", fileName);
        model.addAttribute("sheetName", sheetName);
        model.addAttribute("headers", headers);
        model.addAttribute("data", data);
        return "excel/view-sheet";
    }

    /**
     * Data analysis page
     */
    @GetMapping("/analyze")
    public String analyzeData(Model model) {
        List<File> files = excelService.getExcelFiles();
        List<ExcelFileInfo> fileInfoList = convertToFileInfoList(files);
        model.addAttribute("files", fileInfoList);
        return "excel/analyze";
    }

    /**
     * Process data analysis
     */
    @PostMapping("/analyze")
    public String processAnalysis(@RequestParam String fileName,
                                 @RequestParam(required = false) String sheetName,
                                 @RequestParam String operation,
                                 Model model) {
        
        List<File> files = excelService.getExcelFiles();
        File targetFile = files.stream()
                .filter(file -> file.getName().equals(fileName))
                .findFirst()
                .orElse(null);

        if (targetFile == null) {
            model.addAttribute("error", "File not found: " + fileName);
            return "excel/error";
        }

        List<ExcelData> data;
        if (sheetName != null && !sheetName.trim().isEmpty()) {
            data = excelService.extractSheetData(targetFile, sheetName);
        } else {
            ExcelProcessingResult result = excelService.extractExcelFile(targetFile);
            data = result.getExtractedData();
        }

        Map<String, Object> analysisResult = excelService.performDataOperations(data, operation);
        
        model.addAttribute("fileName", fileName);
        model.addAttribute("sheetName", sheetName);
        model.addAttribute("operation", operation);
        model.addAttribute("result", analysisResult);
        
        return "excel/analysis-result";
    }
}
