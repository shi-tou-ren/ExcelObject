package com.str.newExcel.analysis.config;

import org.apache.poi.hssf.usermodel.HSSFSheet;

public interface ConfigAnalysis {

    /**
     * 初始化表格配置 并返回表格列的映射
     */
    String[] initSheet(HSSFSheet sheet, Object config);

}
