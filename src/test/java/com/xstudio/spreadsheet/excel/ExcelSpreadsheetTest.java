package com.xstudio.spreadsheet.excel;

import com.xstudio.spreadsheet.excel.entity.Header;
import com.xstudio.spreadsheet.excel.entity.Row;
import com.xstudio.spreadsheet.excel.entity.WorkSheet;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

/**
 * @author xiaobiao
 * @version 2020/3/14
 */
class ExcelSpreadsheetTest {

    @Test
    void write() throws IOException {
        ClassLoader classLoader = this.getClass().getClassLoader();
        String path = classLoader.getResource("").getPath();
        String file = path + "demo.xlsx";
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("history_id"));
        headers.add(new Header("tenant_id"));
        List<Row> row = new ArrayList<>();
        Row rowValue;
        for (int i = 0; i < 1000000; i++) {
            rowValue = new Row();
            rowValue.getCells().put("history_id", i);
            rowValue.getCells().put("tenant_id", "CA_PV_TENANT");
            row.add(rowValue);
        }
        long start = System.nanoTime();
        System.out.printf("开始时间 %s", start);
        WorkSheet workSheet = new WorkSheet(headers, row);
        try {
            ExcelSpreadsheet.write(new File(file), workSheet);
        } catch (Exception e) {
            System.out.printf(e.getMessage());
        }
        long end = System.nanoTime();
        System.out.printf("结束时间 %s  耗时 %s", start,  TimeUnit.NANOSECONDS.toSeconds(end - start));
    }
}