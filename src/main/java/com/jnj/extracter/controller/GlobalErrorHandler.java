package com.jnj.extracter.controller;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.servlet.ModelAndView;

import jakarta.servlet.http.HttpServletRequest;

/**
 * Global error handler for the application
 */
@ControllerAdvice
public class GlobalErrorHandler {

    private static final Logger logger = LoggerFactory.getLogger(GlobalErrorHandler.class);
    
    /**
     * Handle all exceptions
     */
    @ExceptionHandler(Exception.class)
    public String handleError(HttpServletRequest request, Exception e, Model model) {
        logger.error("Error handling request: " + request.getRequestURL(), e);
        
        model.addAttribute("error", e.getMessage());
        model.addAttribute("url", request.getRequestURL());
        
        return "excel/error";
    }
}
