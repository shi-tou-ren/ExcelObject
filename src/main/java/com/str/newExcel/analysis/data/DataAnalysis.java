package com.str.newExcel.analysis.data;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import java.util.Collection;

public interface DataAnalysis {

    /**
     * 数据转换为ExcelSheet
     */
    void dataToExcelSheet(HSSFSheet sheet, Object data, String[] columnMapping) throws IllegalAccessException;

}
