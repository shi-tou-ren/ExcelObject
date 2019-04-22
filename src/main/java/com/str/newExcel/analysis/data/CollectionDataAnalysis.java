package com.str.newExcel.analysis.data;

import com.str.newExcel.exception.AnalysisException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.lang.reflect.Field;
import java.util.*;

public class CollectionDataAnalysis implements DataAnalysis {

    public void dataToExcelSheet(HSSFSheet sheet, Object data, String[] columnMapping) throws IllegalAccessException {
        if (!(data instanceof Collection)) {
            throw new AnalysisException("数据类型错误");
        }
        if (!((Collection) data).isEmpty()) {
            if (data instanceof List) {
                listToExcelSheet(sheet, (List) data, columnMapping);
            } else if (data instanceof Set) {
                setToExcelSheet(sheet, (Set) data, columnMapping);
            } else {
                mapToExcelSheet(sheet, (Map) data, columnMapping);
            }
        }
    }

    private void listToExcelSheet(HSSFSheet sheet, List data, String[] columnMapping) {
        Field[] fields = getFileArray(data.get(0).getClass(), columnMapping);
        int index = sheet.getLastRowNum() + 1;
        for (Object obj : data) {
            HSSFRow row = sheet.createRow(index++);
            dataToRow(obj, fields, row);
        }
    }

    private void setToExcelSheet(HSSFSheet sheet, Set data, String[] columnMapping) {
        Field[] fields = null;
        int index = -1;
        for (Object obj : data) {
            if (index == -1) {
                fields = getFileArray(obj.getClass(), columnMapping);
                index = sheet.getLastRowNum() + 1;
            }
            HSSFRow row = sheet.createRow(index++);
            dataToRow(obj, fields, row);
        }
    }

    private void dataToRow(Object date, Field[] fields, HSSFRow row) {
        for (int i = 0; i < fields.length; i++) {
            if (null != fields[i]) {
                try {
                    setCellValue(row.createCell(i), fields[i].get(date));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private void mapToExcelSheet(HSSFSheet sheet, Map data, String[] columnMapping) {
        throw new AnalysisException("暂不支持Map类型的数据解析");
    }

    private Field[] getFileArray(Class clazz, String[] columnMapping) {
        Field[] fields = new Field[columnMapping.length];
        for (int i = 0; i < columnMapping.length; i++) {
            try {
                fields[i] = clazz.getDeclaredField(columnMapping[i]);
                fields[i].setAccessible(true);
            } catch (NoSuchFieldException e) {
                fields[i] = null;
            }
        }
        return fields;
    }

    private void setCellValue(HSSFCell cell, Object cellValue) {
        if (null == cellValue) {
            cell.setCellValue("");
        } else if (cellValue instanceof Number) {
            cell.setCellValue(((Number) cellValue).doubleValue());
        } else {
            cell.setCellValue(cellValue.toString());
        }
    }

}
