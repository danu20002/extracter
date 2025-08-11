package com.jnj.extracter.config;

import com.jnj.extracter.storage.LocalStorageService;
import com.jnj.extracter.storage.SapDmsStorageService;
import com.jnj.extracter.storage.StorageService;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class AttachmentStorageConfig {
    @Value("${attachment.storage.type:local}")
    private String storageType;

    private final LocalStorageService localStorageService;
    private final SapDmsStorageService sapDmsStorageService;

    public AttachmentStorageConfig(LocalStorageService localStorageService, SapDmsStorageService sapDmsStorageService) {
        this.localStorageService = localStorageService;
        this.sapDmsStorageService = sapDmsStorageService;
    }

    @Bean
    public StorageService storageService() {
        if ("sapdms".equalsIgnoreCase(storageType)) {
            return sapDmsStorageService;
        }
        return localStorageService;
    }
}
