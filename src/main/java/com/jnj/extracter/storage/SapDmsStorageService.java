package com.jnj.extracter.storage;

import org.springframework.stereotype.Service;
import java.io.InputStream;

@Service
public class SapDmsStorageService implements StorageService {
    @Override
    public Object save(String fileName, InputStream inputStream) throws Exception {
        // TODO: Implement SAP DMS integration here
        // For now, just return a placeholder string
        return "DMS_ID_PLACEHOLDER";
    }
}
