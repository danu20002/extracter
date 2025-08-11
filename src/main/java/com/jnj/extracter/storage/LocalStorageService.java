package com.jnj.extracter.storage;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;

@Service
public class LocalStorageService implements StorageService {
    @Value("${attachment.save.dir:excel/temp}")
    private String saveDir;

    @Override
    public File save(String fileName, InputStream inputStream) throws Exception {
        File dir = new File(saveDir);
        if (!dir.exists()) dir.mkdirs();
        File file = new File(dir, fileName);
        try (FileOutputStream output = new FileOutputStream(file)) {
            byte[] buffer = new byte[4096];
            int bytesRead;
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                output.write(buffer, 0, bytesRead);
            }
        }
        return file;
    }
}
