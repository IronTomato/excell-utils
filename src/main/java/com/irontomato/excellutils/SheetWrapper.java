package com.irontomato.excellutils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class SheetWrapper {

    private Sheet sheet;

    private int rowCursor = 0;

    private SheetWrapper(Sheet sheet, int rowCursor) {
        this.sheet = sheet;
        this.rowCursor = rowCursor;
    }

    public static SheetWrapper wrap(Sheet sheet) {
        return new SheetWrapper(sheet, 0);
    }

    public static SheetWrapper wrap(Sheet sheet, int startRow) {
        return new SheetWrapper(sheet, startRow);
    }

    public Row addRow(){
        return sheet.createRow(rowCursor++);
    }

    public int mergeRegion(int rowFrom, int colFrom, int rowTo, int colTo) {
        return sheet.addMergedRegion(new CellRangeAddress(rowFrom, colFrom, rowTo, colTo));
    }

    public int mergeRegion(Cell origin, int width, int height) {
        if (width < 1 || height < 1) {
            throw new IllegalArgumentException("width and height should >= 1");
        }
        int rowFrom = origin.getRowIndex();
        int colFrom = origin.getColumnIndex();
        return mergeRegion(rowFrom, colFrom, rowFrom + height -1, colFrom + width -1);
    }

    public int mergeRegion(Cell start, Cell end) {
        return mergeRegion(start.getRowIndex(), start.getColumnIndex(), end.getRowIndex(), end.getColumnIndex());
    }

}
