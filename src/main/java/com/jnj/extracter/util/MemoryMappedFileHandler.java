package com.jnj.extracter.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;

import org.springframework.stereotype.Component;
import lombok.extern.slf4j.Slf4j;

/**
 * Utility class for handling Excel files using memory-mapped I/O for better performance.
 * This provides significant speed improvements for large files by avoiding traditional I/O.
 */
@Component
@Slf4j
public class MemoryMappedFileHandler {

    /**
     * Creates a memory-mapped byte buffer for the given file.
     *
     * @param file The file to map into memory
     * @return A MappedByteBuffer for the file
     * @throws IOException If an I/O error occurs
     */
    public MappedByteBuffer createMemoryMappedBuffer(File file) throws IOException {
        try (RandomAccessFile raf = new RandomAccessFile(file, "r");
             FileChannel channel = raf.getChannel()) {
            
            // Map the entire file into memory
            return channel.map(FileChannel.MapMode.READ_ONLY, 0, channel.size());
        } catch (IOException e) {
            log.error("Error creating memory-mapped buffer for file: {}", file.getName(), e);
            throw e;
        }
    }
    
    /**
     * Reads a file into a direct byte buffer for faster processing.
     * This is an alternative to memory mapping for smaller files.
     *
     * @param file The file to read
     * @param bufferSize The size of the buffer to allocate
     * @return A ByteBuffer containing the file contents
     * @throws IOException If an I/O error occurs
     */
    public ByteBuffer readFileToDirectBuffer(File file, int bufferSize) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             FileChannel channel = fis.getChannel()) {
            
            // Determine the file size and allocate a direct buffer
            long fileSize = channel.size();
            ByteBuffer buffer = ByteBuffer.allocateDirect((int) fileSize);
            
            // Read the file contents into the buffer
            channel.read(buffer);
            buffer.flip();
            
            return buffer;
        } catch (IOException e) {
            log.error("Error reading file to direct buffer: {}", file.getName(), e);
            throw e;
        }
    }
    
    /**
     * Safely releases a MappedByteBuffer if possible.
     * Note: This is a best-effort operation as Java doesn't provide a standard way to release mappings.
     *
     * @param buffer The buffer to release
     */
    public void releaseBuffer(MappedByteBuffer buffer) {
        if (buffer != null) {
            buffer.clear();
            
            // On newer JDKs, this might work to force unmapping
            try {
                if (buffer.getClass().getMethod("cleaner") != null) {
                    Object cleaner = buffer.getClass().getMethod("cleaner").invoke(buffer);
                    if (cleaner != null) {
                        cleaner.getClass().getMethod("clean").invoke(cleaner);
                    }
                }
            } catch (Exception e) {
                // Just ignore if we can't clean - GC will handle it eventually
                log.debug("Could not manually release MappedByteBuffer, will rely on GC");
            }
        }
    }
}
