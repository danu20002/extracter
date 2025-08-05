package com.jnj.extracter.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.Map;

@RestController
public class TestController {

    @GetMapping("/")
    public Map<String, String> home() {
        return Map.of(
            "message", "Excel Data Extractor is running!",
            "version", "1.0.0",
            "endpoints", "/api/excel/*"
        );
    }
}
