package com.xstudio.spreadsheet.excel.entity;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

/**
 * @author xiaobiao
 * @version 2020/3/14
 */
public class WorkSheet {
    @Getter
    @Setter
    private String name = "Sheet1";

    @Getter
    @Setter
    private List<Header> headers;

    @Getter
    @Setter
    private List<Row> rows;

    public WorkSheet(String name, List<Header> headers, List<Row> rows) {
        this.name = name;
        this.headers = headers;
        this.rows = rows;
    }

    public WorkSheet(List<Header> headers, List<Row> rows) {
        this.headers = headers;
        this.rows = rows;
    }
}
