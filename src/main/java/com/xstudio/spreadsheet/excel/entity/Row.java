package com.xstudio.spreadsheet.excel.entity;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

/**
 * @author xiaobiao
 * @version 2020/3/14
 */
@Data
public class Row {
    private Map<String, Object> cells = new HashMap<>();
}
