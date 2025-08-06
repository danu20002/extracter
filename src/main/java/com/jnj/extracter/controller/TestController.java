package com.jnj.extracter.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

@Controller
public class TestController {

    @GetMapping("/")
    @ResponseBody
    public Map<String, String> home() {
        return Map.of(
            "message", "Excel Data Extractor is running!",
            "version", "1.0.0",
            "endpoints", "/api/excel/*"
        );
    }
    
    /**
     * URL test page
     */
    @GetMapping("/test")
    public String testUrls(Model model) {
        // Define all the URLs we want to test
        List<UrlTest> urls = Arrays.asList(
            new UrlTest("/excel", "Dashboard - Main page", "GET"),
            new UrlTest("/excel/files", "Files - List all Excel files", "GET"),
            new UrlTest("/excel/analyze", "Analysis - Configure data analysis", "GET"),
            new UrlTest("/excel/view", "Error - Missing fileName parameter", "GET"),
            new UrlTest("/error", "Error page", "GET")
        );
        
        model.addAttribute("urls", urls);
        return "test/url-tester";
    }
    
    /**
     * URL test result class
     */
    public static class UrlTest {
        private String url;
        private String description;
        private String method;
        private boolean success;
        private String errorMessage;
        
        public UrlTest(String url, String description, String method) {
            this.url = url;
            this.description = description;
            this.method = method;
            this.success = false;
            this.errorMessage = "";
        }
        
        public String getUrl() {
            return url;
        }
        
        public String getDescription() {
            return description;
        }
        
        public String getMethod() {
            return method;
        }
        
        public boolean isSuccess() {
            return success;
        }
        
        public void setSuccess(boolean success) {
            this.success = success;
        }
        
        public String getErrorMessage() {
            return errorMessage;
        }
        
        public void setErrorMessage(String errorMessage) {
            this.errorMessage = errorMessage;
        }
    }
}
