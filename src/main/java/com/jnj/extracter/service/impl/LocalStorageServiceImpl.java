package com.jnj.extracter.service.impl;

import com.jnj.extracter.api.exception.FileStorageException;
import com.jnj.extracter.api.exception.ResourceNotFoundException;
import com.jnj.extracter.config.StorageProperties;
import com.jnj.extracter.service.contract.StorageService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Service;
import org.springframework.util.FileSystemUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.List;
import java.util.UUID;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Implementation of StorageService that stores files in the local file system.
 */
@Service("localStorageService")
@Slf4j
public class LocalStorageServiceImpl implements StorageService {

    private final Path rootLocation;

    public LocalStorageServiceImpl(StorageProperties properties) {
        this.rootLocation = Paths.get(properties.getLocation());
        init();
        log.info("LocalStorageServiceImpl initialized with root location: {}", rootLocation);
    }

    @Override
    public void init() {
        try {
            Files.createDirectories(rootLocation);
            log.debug("Created directories at {}", rootLocation);
        } catch (IOException e) {
            log.error("Could not initialize storage location", e);
            throw new FileStorageException("Could not initialize storage location", e);
        }
    }

    @Override
    public String store(MultipartFile file) {
        try {
            if (file.isEmpty()) {
                log.warn("Failed to store empty file");
                throw new FileStorageException("Failed to store empty file");
            }
            
            String filename = generateUniqueFilename(file.getOriginalFilename());
            Path destinationFile = this.rootLocation.resolve(Paths.get(filename)).normalize().toAbsolutePath();
            
            if (!destinationFile.getParent().equals(this.rootLocation.toAbsolutePath())) {
                log.warn("Cannot store file outside current directory");
                throw new FileStorageException("Cannot store file outside current directory");
            }
            
            try (InputStream inputStream = file.getInputStream()) {
                Files.copy(inputStream, destinationFile, StandardCopyOption.REPLACE_EXISTING);
                log.info("Stored file: {} (size: {})", filename, file.getSize());
                return filename;
            }
        } catch (IOException e) {
            log.error("Failed to store file", e);
            throw new FileStorageException("Failed to store file", e);
        }
    }
    
    @Override
    public String store(File file) throws IOException {
        try {
            if (!file.exists()) {
                log.warn("Failed to store non-existent file: {}", file.getAbsolutePath());
                throw new FileStorageException("Failed to store non-existent file");
            }
            
            String filename = generateUniqueFilename(file.getName());
            Path destinationFile = this.rootLocation.resolve(Paths.get(filename)).normalize().toAbsolutePath();
            
            if (!destinationFile.getParent().equals(this.rootLocation.toAbsolutePath())) {
                log.warn("Cannot store file outside current directory");
                throw new FileStorageException("Cannot store file outside current directory");
            }
            
            Files.copy(file.toPath(), destinationFile, StandardCopyOption.REPLACE_EXISTING);
            log.info("Stored file: {} (size: {})", filename, file.length());
            return filename;
        } catch (IOException e) {
            log.error("Failed to store file", e);
            throw new FileStorageException("Failed to store file", e);
        }
    }

    @Override
    public List<Path> loadAll() throws IOException {
        try (Stream<Path> stream = Files.walk(this.rootLocation, 1)) {
            return stream
                    .filter(path -> !path.equals(this.rootLocation))
                    .map(this.rootLocation::relativize)
                    .map(this.rootLocation::resolve)
                    .collect(Collectors.toList());
        } catch (IOException e) {
            log.error("Failed to read stored files", e);
            throw new FileStorageException("Failed to read stored files", e);
        }
    }

    @Override
    public Path load(String filename) {
        Path file = rootLocation.resolve(filename);
        if (Files.exists(file)) {
            return file;
        }
        log.warn("File not found: {}", filename);
        return null;
    }

    @Override
    public Resource loadAsResource(String filename) {
        try {
            Path file = load(filename);
            if (file == null) {
                throw new ResourceNotFoundException("File not found: " + filename);
            }
            
            Resource resource = new UrlResource(file.toUri());
            if (resource.exists() || resource.isReadable()) {
                return resource;
            } else {
                log.warn("Could not read file: {}", filename);
                throw new FileStorageException("Could not read file: " + filename);
            }
        } catch (MalformedURLException e) {
            log.error("Could not read file: {}", filename, e);
            throw new FileStorageException("Could not read file: " + filename, e);
        }
    }

    @Override
    public void delete(String filename) {
        Path file = load(filename);
        if (file != null) {
            try {
                Files.deleteIfExists(file);
                log.info("Deleted file: {}", filename);
            } catch (IOException e) {
                log.error("Could not delete file: {}", filename, e);
                throw new FileStorageException("Could not delete file: " + filename, e);
            }
        }
    }

    @Override
    public void deleteAll() {
        FileSystemUtils.deleteRecursively(rootLocation.toFile());
        log.info("Deleted all files in {}", rootLocation);
        init();
    }
    
    @Override
    public long getFileSize(Path path) throws IOException {
        return Files.size(path);
    }
    
    @Override
    public LocalDateTime getLastModifiedTime(Path path) throws IOException {
        return LocalDateTime.ofInstant(
                Files.getLastModifiedTime(path).toInstant(),
                ZoneId.systemDefault()
        );
    }
    
    @Override
    public Path getRootLocation() {
        return rootLocation;
    }
    
    private String generateUniqueFilename(String originalFilename) {
        String extension = "";
        String name = originalFilename;
        
        int lastDotIndex = originalFilename.lastIndexOf('.');
        if (lastDotIndex > 0) {
            extension = originalFilename.substring(lastDotIndex);
            name = originalFilename.substring(0, lastDotIndex);
        }
        
        String uniqueId = UUID.randomUUID().toString().substring(0, 8);
        return name + "_" + uniqueId + extension;
    }
}
