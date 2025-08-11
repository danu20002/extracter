package com.jnj.extracter.util;

import java.util.ArrayList;
import java.util.List;

import org.springframework.stereotype.Component;

import com.jnj.extracter.domain.model.ExcelData;
import com.jnj.extracter.domain.model.ExcelProcessingResult;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;

import lombok.extern.slf4j.Slf4j;

/**
 * Utility class to convert between entity objects and JSON for data transfer.
 * This class replaces the Protocol Buffer implementation with a JSON-based approach
 * for better compatibility.
 */
@Component
@Slf4j
public class ProtoConverter {
    
    private final ObjectMapper objectMapper = new ObjectMapper();
    
    /**
     * Convert an entity ExcelProcessingResult to a JSON representation.
     *
     * @param result The entity object
     * @return JsonNode representing the data
     */
    public JsonNode toJson(ExcelProcessingResult result) {
        ObjectNode jsonResult = objectMapper.createObjectNode()
                .put("filename", result.getFilename())
                .put("success", result.isSuccess())
                .put("message", result.getMessage() != null ? result.getMessage() : "")
                .put("sheetCount", result.getSheetCount())
                .put("totalRowCount", result.getTotalRowCount());
        
        // Add sheet names
        if (result.getSheetNames() != null) {
            jsonResult.set("sheetNames", objectMapper.valueToTree(result.getSheetNames()));
        }
        
        // Convert and add extracted data
        if (result.getExtractedData() != null && !result.getExtractedData().isEmpty()) {
            jsonResult.set("extractedData", objectMapper.valueToTree(result.getExtractedData()));
        }
        
        return jsonResult;
    }
    
    /**
     * Convert a JSON representation to an entity ExcelProcessingResult.
     *
     * @param json The JSON node
     * @return The entity object
     */
    public ExcelProcessingResult fromJson(JsonNode json) {
        ExcelProcessingResult result = new ExcelProcessingResult();
        result.setFilename(json.path("filename").asText(""));
        result.setSuccess(json.path("success").asBoolean());
        result.setMessage(json.path("message").asText(""));
        result.setSheetCount(json.path("sheetCount").asInt());
        result.setTotalRowCount(json.path("totalRowCount").asInt());
        
        // Convert sheet names
        if (json.has("sheetNames") && json.get("sheetNames").isArray()) {
            List<String> sheetNames = new ArrayList<>();
            for (JsonNode name : json.get("sheetNames")) {
                sheetNames.add(name.asText());
            }
            result.setSheetNames(sheetNames);
        }
        
        // Convert extracted data
        if (json.has("extractedData") && json.get("extractedData").isArray()) {
            List<ExcelData> extractedData = new ArrayList<>();
            
            for (JsonNode row : json.get("extractedData")) {
                extractedData.add(jsonNodeToExcelData(row));
            }
            
            result.setExtractedData(extractedData);
        }
        
        return result;
    }
    
    /**
     * Convert an entity ExcelData to a JSON node.
     *
     * @param data The entity object
     * @return The JSON node
     */
    public JsonNode excelDataToJsonNode(ExcelData data) {
        ObjectNode rowNode = objectMapper.createObjectNode()
            .put("filename", data.getFilename())
            .put("sheetCount", data.getSheetCount());
        
        // Convert sheets data
        if (data.getSheets() != null) {
            rowNode.set("sheets", objectMapper.valueToTree(data.getSheets()));
        }
        
        return rowNode;
    }
    
    /**
     * Convert a JSON node to an entity ExcelData.
     *
     * @param row The JSON node
     * @return The entity object
     */
    public ExcelData jsonNodeToExcelData(JsonNode row) {
        ExcelData data = new ExcelData();
        data.setFilename(row.path("filename").asText(""));
        data.setSheetCount(row.path("sheetCount").asInt());
        
        // Convert sheets
        if (row.has("sheets") && row.get("sheets").isArray()) {
            // Assuming there's an appropriate method to deserialize sheets
            // This would require proper implementation based on your SheetData structure
            // For now, we'll just log that sheets data exists
            log.debug("Sheet data found in JSON but not processed in this implementation");
        }
        
        return data;
    }
    
    // Legacy method names to maintain compatibility with existing code
    
    /**
     * Legacy method for backward compatibility. 
     * @param result The entity object
     * @return A String representation in JSON format
     */
    public String toProto(ExcelProcessingResult result) {
        try {
            return objectMapper.writeValueAsString(toJson(result));
        } catch (Exception e) {
            log.error("Error serializing result to JSON", e);
            return "{}";
        }
    }
    
    /**
     * Legacy method for backward compatibility.
     * @param data The Excel data object
     * @return A String representation in JSON format
     */
    public String toProto(ExcelData data) {
        try {
            return objectMapper.writeValueAsString(excelDataToJsonNode(data));
        } catch (Exception e) {
            log.error("Error serializing data to JSON", e);
            return "{}";
        }
    }
    
    /**
     * Legacy method for backward compatibility.
     * @param json The JSON string
     * @return The ExcelProcessingResult object
     */
    public ExcelProcessingResult fromProto(String json) {
        try {
            return fromJson(objectMapper.readTree(json));
        } catch (Exception e) {
            log.error("Error deserializing JSON to result", e);
            return new ExcelProcessingResult();
        }
    }
    
    /**
     * Legacy method for backward compatibility.
     * @param json The JSON string
     * @return The ExcelData object
     */
    public ExcelData fromProto(String json, boolean isRow) {
        try {
            return jsonNodeToExcelData(objectMapper.readTree(json));
        } catch (Exception e) {
            log.error("Error deserializing JSON to data", e);
            return new ExcelData();
        }
    }
}
