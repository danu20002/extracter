package com.jnj.extracter.domain.model;

import lombok.Data;

/**
 * Represents a cell of data in an Excel sheet.
 */
@Data
public class CellData {
    private int rowIndex;
    private int columnIndex;
    private String header;
    private String value;
    private String type;
    private String formula;
}
