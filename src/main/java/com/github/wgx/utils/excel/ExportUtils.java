package com.github.wgx.utils.excel;

import static com.google.common.collect.Lists.newArrayList;
import static java.util.stream.Collectors.toMap;

import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Function;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.phantomthief.util.MoreFunctions;
import com.github.wgx.utils.json.ObjectMapperUtils;
import com.google.common.collect.ImmutableList;

/**
 * @author derek.w
 * Created on 2021-06-11
 */
@SuppressWarnings("rawtypes")
public class ExportUtils {
    private static final Logger log = LoggerFactory.getLogger(ExportUtils.class);

    public static <T> byte[] exportExcel(List<T> dataList, Class<T> clazz, ExcelType excelType) {
        checkExcelType(excelType);
        Workbook workbook = export2Excel(dataList, clazz, excelType);
        return writeWorkBook2Bytes(workbook);
    }


    public static <T> byte[] exportExcel(List<T> dataList, Class<T> clazz,
            List<ExcelFileColumn> headers, ExcelType excelType) {
        checkExcelType(excelType);
        Workbook workbook = export2Excel(dataList, clazz, headers, excelType);
        return writeWorkBook2Bytes(workbook);
    }

    private static byte[] writeWorkBook2Bytes(Workbook workbook) {
        byte[] bytes = null;
        try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
            workbook.write(os);
            bytes = os.toByteArray();
        } catch (Exception e) {
            log.error("", e);
        }
        return bytes;
    }

    private static void checkExcelType(ExcelType excelType) {
        if (excelType == null) {
            throw new IllegalArgumentException("unknown excel type");
        }
    }


    // annotation based
    public static <T> Workbook export2Excel(List<T> dataList, Class<T> clazz, ExcelType excelType) {
        checkExcelType(excelType);
        Workbook workbook = createWorkBook(excelType);
        Sheet sheet = workbook.createSheet();
        Map<Class, List<Field>> allClassAndField = getAllFields(clazz);
        if (MapUtils.isEmpty(allClassAndField)) {
            throw new RuntimeException("no classes or fields found");
        }
        List<Pair<String, Integer>> headers = getAllHeaders(allClassAndField);

        writeHeaderByAnnotation(sheet, headers);
        writeBodyByAnnotation(sheet, dataList, clazz);
        return workbook;
    }

    // specify columns
    public static <T> Workbook export2Excel(List<T> dataList, Class<T> clazz,
            List<ExcelFileColumn> columns, ExcelType excelType) {
        checkExcelType(excelType);
        Workbook workbook = createWorkBook(excelType);
        Sheet sheet = workbook.createSheet();
        writeSpecifiedHeader(sheet, columns);
        Map<String, ExcelFileColumn> columnMap = columns.stream().collect(toMap(ExcelFileColumn::getObjField, Function.identity()));
        writeBodyByHeader(sheet, dataList, clazz, columnMap);
        return workbook;
    }

    private static <T> void writeBodyByHeader(Sheet sheet, List<T> dataList,
            Class<T> clazz, Map<String, ExcelFileColumn> columnMap) {
        int index = 1;
        for (T data : dataList) {
            Row row =  sheet.createRow(index);
            MoreFunctions.runCatching(() -> writeRowByColumn(row, data, clazz, columnMap));
            index++;
        }
    }

    private static void writeSpecifiedHeader(Sheet sheet, List<ExcelFileColumn> headers) {
        Row headerRow = sheet.createRow(0);
        for (ExcelFileColumn header : headers) {
            Cell cell = headerRow.createCell(header.getColumnIndex());
            cell.setCellValue(header.getHeader());
        }
    }


    private static Workbook createWorkBook(ExcelType excelType) {
        return excelType == ExcelType.XLSX
               ? new XSSFWorkbook()
               : new HSSFWorkbook();
    }


    private static void writeHeaderByAnnotation(Sheet sheet, List<Pair<String, Integer>> headers) {
        Row headerRow = sheet.createRow(0);
        for (Pair<String, Integer> pair : headers) {
            Cell cell = headerRow.createCell(pair.getRight());
            cell.setCellValue(pair.getLeft());
        }
    }

    private static <T> void writeBodyByAnnotation(Sheet sheet, List<T> dataList, Class<T> clazz) {
        int index = 1;
        for (T data : dataList) {
            Row row =  sheet.createRow(index);
            MoreFunctions.runCatching(() -> writeRowByAnnotation(row, data, clazz));
            index++;
        }
    }

    private static <T> void writeRowByAnnotation(Row row, Object data, Class<T> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        if (data == null || fields == null || fields.length == 0) {
            return;
        }
        if (isSimpleClass(clazz.getSuperclass())) {
            writeRowByAnnotation(row, data, clazz.getSuperclass());
        }
        for (Field field : fields) {
            field.setAccessible(true);
            ExcelColumn column = MoreFunctions.catching(() -> field.getAnnotation(ExcelColumn.class));
            if (column != null) {
                Cell cell = row.createCell(column.columnIndex());
                // 直接写column
                if (writeCell(field, cell, data, null)) {
                    continue;
                }
            }
            if (isSimpleClass(field.getType())) {
                writeRowByAnnotation(row, MoreFunctions.catching(() -> field.get(data)), field.getType());
            }
        }
    }

    private static <T> void writeRowByColumn(Row row, Object data, Class<T> clazz,
            Map<String, ExcelFileColumn> headerMap) {
        Field[] fields = clazz.getDeclaredFields();
        if (data == null || fields == null || fields.length == 0) {
            return;
        }
        if (isSimpleClass(clazz.getSuperclass())) {
            writeRowByColumn(row, data, clazz.getSuperclass(), headerMap);
        }
        for (Field field : fields) {
            field.setAccessible(true);
            ExcelFileColumn fileHeader = headerMap.get(field.getName());
            if (fileHeader != null) {
                Cell cell = row.createCell(fileHeader.getColumnIndex());
                // 直接写column
                if (writeCell(field, cell, data, fileHeader.getFormatter())) {
                    continue;
                }
            }
            if (field.getType() != clazz && isSimpleClass(field.getType())) {
                writeRowByColumn(row, MoreFunctions.catching(() -> field.get(data)), field.getType(), headerMap);
            }
        }
    }

    private static boolean writeCell(Field field, Cell cell, Object data, Function<Object, String> func) {
        String fieldTypeName = field.getType().getSimpleName();
        if (isNumericField(fieldTypeName)) {
            Number number = MoreFunctions.catching(() -> (Number) field.get(data));
            cell.setCellValue(number == null ? "NAN" : (func == null ? number.toString() : func.apply(number)));
            return true;
        }
        if (isStringField(fieldTypeName)) {
            cell.setCellValue(MoreFunctions.catching(() -> (String) field.get(data)));
            return true;
        }
        if (isBooleanFiled(fieldTypeName)) {
            cell.setCellValue(MoreFunctions.catching(() -> (Boolean) field.get(data)));
            return true;
        }
        if (isCollectionField(fieldTypeName) || isMapFiled(fieldTypeName)) {
            Object obj = MoreFunctions.catching(() -> field.get(data));
            cell.setCellValue(obj == null ? "" : ObjectMapperUtils.toJSON(obj));
            return true;
        }
        return false;
    }


    private static <T> Map<Class, List<Field>> getAllFields(Class<T> clazz) {
        Map<Class, List<Field>> allRelatedClasses = new HashMap<>();
        Class curClazz = clazz;
        //解决继承问题
        while (curClazz != null && !curClazz.getName().toLowerCase().equals("java.lang.object")) {
            allRelatedClasses.put(curClazz, ImmutableList.copyOf(curClazz.getDeclaredFields()));
            curClazz = curClazz.getSuperclass();
        }
        // 解决某个类中引用其他类的问题
        Map<Class, List<Field>> temp = new HashMap<>();
        for (Class c : allRelatedClasses.keySet()) {
            temp.putAll(processClassFiled(c));
        }
        allRelatedClasses.putAll(temp);
        return allRelatedClasses;
    }

    private static List<Pair<String, Integer>> getAllHeaders(Map<Class, List<Field>> allFieldsWithClass) {
        List<Pair<String, Integer>> headers = newArrayList();
        for (Map.Entry<Class, List<Field>> entry : allFieldsWithClass.entrySet()) {
            List<Field> fields = entry.getValue();
            if (CollectionUtils.isEmpty(fields)) {
                continue;
            }
            for (Field field : fields) {
                ExcelColumn column = MoreFunctions.catching(() -> field.getAnnotation(ExcelColumn.class));
                if (column == null) {
                    continue;
                }
                headers.add(Pair.of(column.header(), column.columnIndex()));
            }
        }
        return headers;
    }

    private static Map<Class, List<Field>> processClassFiled(Class c) {
        Field[] fields = c.getDeclaredFields();
        if (fields == null || fields.length == 0) {
            return Collections.emptyMap();
        }
        Map<Class, List<Field>> classAndField = new HashMap<>();
        classAndField.put(c, ImmutableList.copyOf(fields));
        for (Field field : fields) {
            field.setAccessible(true);
            String fieldTypeName = field.getType().getSimpleName();
            // 直接写column
            if (isSimpleField(fieldTypeName)
                    || isMapFiled(fieldTypeName)
                    || isCollectionField(fieldTypeName)) {
                continue;
            }
            if (field.getType() != c && isSimpleClass(field.getType())) {
                classAndField.putAll(processClassFiled(field.getType()));
            }
        }
        return classAndField;
    }

    // 这里不考虑char及其封装类
    private static boolean isSimpleField(String fieldTypeName) {
        return isStringField(fieldTypeName) || isNumericField(fieldTypeName) || isBooleanFiled(fieldTypeName);
    }

    private static boolean isStringField(String fieldTypeName) {
        return String.class.getSimpleName().equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isNumericField(String fieldTypeName) {
        return isByteFiled(fieldTypeName) || isShortFiled(fieldTypeName) || isIntFiled(fieldTypeName)
                || isFloatFiled(fieldTypeName) || isLongField(fieldTypeName) || isDoubleFiled(fieldTypeName);
    }

    private static boolean isByteFiled(String fieldTypeName) {
        return Byte.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "byte".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isShortFiled(String fieldTypeName) {
        return Short.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "short".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isIntFiled(String fieldTypeName) {
        return Integer.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "int".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isFloatFiled(String fieldTypeName) {
        return Float.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "float".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isLongField(String fieldTypeName) {
        return Long.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "long".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isDoubleFiled(String fieldTypeName) {
        return Double.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "double".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isBooleanFiled(String fieldTypeName) {
        return Boolean.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || "boolean".equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isCollectionField(String fieldTypeName) {
        return List.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || Set.class.getSimpleName().equalsIgnoreCase(fieldTypeName)
                || Collection.class.getSimpleName().equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isMapFiled(String fieldTypeName) {
        return Map.class.getSimpleName().equalsIgnoreCase(fieldTypeName);
    }

    private static boolean isSimpleClass(Class c) {
        return c != null && !Object.class.getName().equalsIgnoreCase(c.getName());
    }
}
