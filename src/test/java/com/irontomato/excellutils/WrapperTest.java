package com.irontomato.excellutils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class WrapperTest {

    @Test
    public void test(){
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("test sheet");
        SheetWrapper wrapper = SheetWrapper.wrap(sheet);
        Row row = wrapper.addRow();
        RowWrapper rowWrapper = RowWrapper.wrap(row);
        rowWrapper.addCells("one", "two", null, null, "three");
    }
}
