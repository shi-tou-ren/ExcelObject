package com.str.newExcel.analysis;

import com.str.newExcel.analysis.config.ClassConfigAnalysis;
import com.str.newExcel.analysis.config.ConfigAnalysis;
import com.str.newExcel.analysis.data.CollectionDataAnalysis;
import com.str.newExcel.analysis.data.DataAnalysis;
import com.str.newExcel.analysis.data.IOStreamAnalysis;
import com.str.newExcel.analysis.data.StreamAnalysis;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;

/**
 * 使用外观模式的解析类 包含 各种数据解析为excel表格和excel表格解析为其他数据
 */
public class AnalysisFacade {

    public static String[] initSheet(HSSFSheet sheet, Object config) {
        return getConfigAnalysis(config).initSheet(sheet, config);
    }

    /**
     *  将数据解析根据配置解析为excel文件中的表格
     */
    public static void dataToExcelSheet(HSSFSheet sheet, String[] columnMapping, Object data) throws IllegalAccessException {
        getDataAnalysis(data).dataToExcelSheet(sheet, data, columnMapping);
    }

    /**
     *  将数据解析为excel文件
     */
    public static HSSFWorkbook dataToExcel(Object data) {
        return null;
    }


    public static void ExcelToStream(Object outPut, HSSFWorkbook excel) throws IOException {
        getStreamAnalysis(outPut).excelToSteam(outPut, excel);
    }

    private static ConfigAnalysis getConfigAnalysis(Object config) {
        return new ClassConfigAnalysis();
    }

    private static DataAnalysis getDataAnalysis(Object data) {
        return new CollectionDataAnalysis();
    }

    private static StreamAnalysis getStreamAnalysis(Object outPut) {
        return new IOStreamAnalysis();
    }
}
