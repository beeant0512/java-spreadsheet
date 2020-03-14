package com.xstudio.spreadsheet.excel.entity;

import lombok.Getter;
import lombok.Setter;

/**
 * @author xiaobiao
 * @version 2020/3/14
 */
public class Header {
    @Getter
    @Setter
    private int width = 110;

    @Getter
    @Setter
    private String value;

    @Getter
    @Setter
    private String key;

    @Getter
    @Setter
    private CellType type = CellType.String;

    public Header(String key) {
        this.value = key;
        this.key = key;
    }

    public Header(String value, String key) {
        this.value = value;
        this.key = key;
    }

    public String getColumn(int index) {
        return "<Column ss:Index=\"" + index + "\" ss:AutoFitWidth=\"0\" ss:Width=\"" + this.getWidth() + "\"/>";
    }

    public String getCell() {
        return "<Cell><Data ss:Type=\"" + getType().name() + "\">" + getValue() + "</Data></Cell>";
    }
}
