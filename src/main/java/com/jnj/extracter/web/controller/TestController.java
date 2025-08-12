package com.jnj.extracter.web.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Controller for testing application endpoints.
 */
@Controller
@RequestMapping("/test")
public class TestController {

    /**
     * Page for testing all endpoints in the application.
     * 
     * @param model The model for the view
     * @return The test page
     */
    @GetMapping("/endpoints")
    public String testEndpoints(Model model) {
        // Create a list of endpoints to test
        List<Map<String, String>> urls = new ArrayList<>();
        
        // Web UI Endpoints
        urls.add(Map.of(
            "url", "/excel/dashboard", 
            "description", "Main dashboard page", 
            "method", "GET"
        ));
        
        urls.add(Map.of(
            "url", "/excel/files", 
            "description", "List all Excel files", 
            "method", "GET"
        ));
        
        urls.add(Map.of(
            "url", "/excel/file/Journal.xlsx", 
            "description", "View specific Excel file", 
            "method", "GET"
        ));
        
        urls.add(Map.of(
            "url", "/excel/upload", 
            "description", "Upload Excel file (requires form data)", 
            "method", "POST"
        ));
        
        // API Endpoints
        urls.add(Map.of(
            "url", "/api/excel/sheets/Journal.xlsx", 
            "description", "Get sheet names for a file", 
            "method", "GET"
        ));
        
        urls.add(Map.of(
            "url", "/api/transform/combine?sourceColumns=Column1,Column2&targetColumn=Combined&separator=%20&fileName=transformed.xlsx", 
            "description", "Transform by combining columns", 
            "method", "POST"
        ));
        
        urls.add(Map.of(
            "url", "/api/transform/download/transformed.xlsx", 
            "description", "Download transformed file", 
            "method", "GET"
        ));
        
        urls.add(Map.of(
            "url", "/api/stats", 
            "description", "Get Excel processing statistics", 
            "method", "GET"
        ));
        
        // Add the URLs to the model
        model.addAttribute("urls", urls);
        model.addAttribute("title", "URL Tester");
        
        return "test/url-tester";
    }
}
