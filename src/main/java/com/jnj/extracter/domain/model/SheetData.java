package com.jnj.extracter.domain.model;

import lombok.Data;

import java.util.List;

/**
 * Represents the data in a single Excel sheet.
 */
@Data
public class SheetData {
    private String name;
    private int index;
    private List<String> headers;
    private List<RowData> rows;
}
