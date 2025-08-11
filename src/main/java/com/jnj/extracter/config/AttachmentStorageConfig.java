package com.jnj.extracter.config;

import com.jnj.extracter.service.contract.StorageService;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class AttachmentStorageConfig {
    @Value("${attachment.storage.type:local}")
    private String storageType;

    private final StorageService localStorageService;
    private final StorageService sapDmsStorageService;

    public AttachmentStorageConfig(
            @Qualifier("localStorageService") StorageService localStorageService, 
            @Qualifier("sapDmsStorageService") StorageService sapDmsStorageService) {
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
