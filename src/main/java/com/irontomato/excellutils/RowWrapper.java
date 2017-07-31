package com.irontomato.excellutils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

public class RowWrapper {

    private Row row;

    private int colCursor = 0;

    private CellStyle defaultStyle;

    private RowWrapper(Row row, int colCursor, CellStyle defaultStyle) {
        this.row = row;
        this.colCursor = colCursor;
        this.defaultStyle = defaultStyle;
    }

    public static RowWrapper wrap(Row row) {
        return new RowWrapper(row, 0, null);
    }

    public static RowWrapper wrap(Row row, int startColumn) {
        return new RowWrapper(row, startColumn, null);
    }

    public static RowWrapper wrap(Row row, CellStyle style) {
        return new RowWrapper(row, row.getLastCellNum(), style);
    }

    public int getColCursor() {
        return colCursor;
    }

    public void setColCursor(int colCursor) {
        this.colCursor = colCursor;
    }

    public void setDefaultStyle(CellStyle defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public Cell addCell(String value, CellStyle style){
        Cell cell = row.createCell(colCursor++);
        if (value != null) {
            cell.setCellValue(value);
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
        return cell;
    }

    public Cell addCell(String value) {
        return addCell(value, defaultStyle);
    }

    public Cell addCell() {
        return addCell(null, defaultStyle);
    }

    public Cell[] addCells(String... values) {
        if (values == null || values.length == 0) {
            throw new IllegalArgumentException("values should not empty");
        }
        Cell[] cells = new Cell[values.length];
        for (int i = 0; i < values.length; i++) {
            cells[i] = addCell(values[i]);
        }
        return cells;
    }

}
