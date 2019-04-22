package com.str.newExcel.analysis.config;

import com.str.newExcel.annotation.ColumnConf;
import com.str.newExcel.annotation.SheetConf;
import com.str.newExcel.exception.AnalysisException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.lang.reflect.Field;

public class ClassConfigAnalysis implements ConfigAnalysis{


    public String[] initSheet(HSSFSheet sheet, Object config) {
        if (!(config instanceof Class)) {
            throw new AnalysisException("配置文件类型出错误，请使用Class类型的配置文件");
        }
        Class clazz = (Class) config;
        SheetConf sheetConf = (SheetConf) clazz.getAnnotation(SheetConf.class);
        initSheetName(sheet, null == sheetConf ||  sheetConf.sheetName().isEmpty() ? clazz.getSimpleName() : sheetConf.sheetName());
        return initSheetTitle(sheet, clazz, null == sheetConf || sheetConf.isHaveTitle());

    }

    private void initSheetName(HSSFSheet sheet, String sheetName) {
        sheet.getWorkbook().setSheetName(sheet.getWorkbook().getSheetIndex(sheet), sheetName);
    }

    private String[] initSheetTitle(HSSFSheet sheet, Class clazz, boolean isHaveTitle) {
        Field[] fields = clazz.getDeclaredFields();
        String[] columnMapping = new String[fields.length];
        HSSFRow row = null;
        if (isHaveTitle) {
            row = sheet.createRow(0);
        }
        for (int i = 0; i < fields.length; i++) {
            if (isHaveTitle) {
                ColumnConf columnConf = fields[i].getAnnotation(ColumnConf.class);
                if (null == columnConf || columnConf.titleName().isEmpty()) {
                    row.createCell(i).setCellValue(fields[i].getName());
                } else {
                    row.createCell(i).setCellValue(columnConf.titleName());
                }
            }
            columnMapping[i] = fields[i].getName();
        }
        return columnMapping;
    }

}
