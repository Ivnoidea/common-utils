package com.github.wgx.utils.excel;

import java.util.function.Function;

import org.apache.commons.lang3.StringUtils;


/**
 * @author derek.w
 * Created on 2021-06-11
 */
public class ExcelFileColumn {
    private final int columnIndex;
    private final String objField;
    private final String header;
    private final Function<Object, String> formatter;

    public ExcelFileColumn(int columnIndex, String objField, String header,
            Function<Object, String> formatter) {
        this.columnIndex = columnIndex;
        this.objField = objField;
        this.header = header;
        this.formatter = formatter;
    }

    public static ExcelFileColumn of(int columnIndex, String objField, String header) {
        return of(columnIndex, objField, header, null);
    }

    public static ExcelFileColumn of(int columnIndex, String objField, String header, Function<Object, String> func) {
        if (columnIndex < 0) {
            throw new IllegalArgumentException("column index is invalid");
        }
        if (StringUtils.isEmpty(objField)) {
            throw new IllegalArgumentException("objField is invalid");
        }
        return new ExcelFileColumn(columnIndex, objField, header, func);
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public String getObjField() {
        return objField;
    }

    public String getHeader() {
        return header;
    }

    public Function<Object, String> getFormatter() {
        return formatter;
    }
}
