package com.jnj.extracter.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Configuration;

/**
 * Configuration properties for email connection.
 */
@Configuration
@ConfigurationProperties(prefix = "application.email")
@Data
public class EmailConfig {
    private String host;
    private int port;
    private String username;
    private String password;
    private String protocol = "imaps";
    private boolean enableSsl = true;
    private String folder = "INBOX";
}
