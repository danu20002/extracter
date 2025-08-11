package com.jnj.extracter.domain.model;

import lombok.Data;

import java.util.List;

/**
 * Represents a row of data in an Excel sheet.
 */
@Data
public class RowData {
    private int index;
    private List<CellData> cells;
    private boolean isHeader;
}
