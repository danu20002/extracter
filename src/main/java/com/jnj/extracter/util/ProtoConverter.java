package com.jnj.extracter.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Component;

import com.jnj.extracter.entity.ExcelData;
import com.jnj.extracter.entity.ExcelProcessingResult;
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
                .put("fileName", result.getFileName())
                .put("success", result.isSuccess())
                .put("message", result.getMessage() != null ? result.getMessage() : "")
                .put("totalSheets", result.getTotalSheets())
                .put("totalRows", result.getTotalRows());
        
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
        result.setFileName(json.path("fileName").asText(""));
        result.setSuccess(json.path("success").asBoolean());
        result.setMessage(json.path("message").asText(""));
        result.setTotalSheets(json.path("totalSheets").asInt());
        result.setTotalRows(json.path("totalRows").asInt());
        
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
            .put("fileName", data.getFileName())
            .put("sheetName", data.getSheetName())
            .put("rowNumber", data.getRowNumber())
            .put("extractedAt", data.getExtractedAt() != null ? data.getExtractedAt() : "");
        
        // Convert cell values
        if (data.getData() != null) {
            ObjectNode dataNode = objectMapper.createObjectNode();
            for (Map.Entry<String, Object> entry : data.getData().entrySet()) {
                if (entry.getValue() != null) {
                    if (entry.getValue() instanceof String) {
                        dataNode.put(entry.getKey(), (String) entry.getValue());
                    } else if (entry.getValue() instanceof Number) {
                        dataNode.put(entry.getKey(), ((Number) entry.getValue()).doubleValue());
                    } else if (entry.getValue() instanceof Boolean) {
                        dataNode.put(entry.getKey(), (Boolean) entry.getValue());
                    } else {
                        dataNode.put(entry.getKey(), entry.getValue().toString());
                    }
                } else {
                    dataNode.putNull(entry.getKey());
                }
            }
            rowNode.set("data", dataNode);
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
        data.setFileName(row.path("fileName").asText(""));
        data.setSheetName(row.path("sheetName").asText(""));
        data.setRowNumber(row.path("rowNumber").asInt());
        data.setExtractedAt(row.path("extractedAt").asText(""));
        
        // Convert data map
        if (row.has("data") && row.get("data").isObject()) {
            Map<String, Object> dataMap = new HashMap<>();
            JsonNode dataNode = row.get("data");
            
            Iterator<String> fieldNames = dataNode.fieldNames();
            while (fieldNames.hasNext()) {
                String fieldName = fieldNames.next();
                JsonNode value = dataNode.get(fieldName);
                
                if (value.isTextual()) {
                    dataMap.put(fieldName, value.asText());
                } else if (value.isNumber()) {
                    dataMap.put(fieldName, value.asDouble());
                } else if (value.isBoolean()) {
                    dataMap.put(fieldName, value.asBoolean());
                } else if (value.isNull()) {
                    dataMap.put(fieldName, null);
                } else {
                    dataMap.put(fieldName, value.toString());
                }
            }
            data.setData(dataMap);
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
