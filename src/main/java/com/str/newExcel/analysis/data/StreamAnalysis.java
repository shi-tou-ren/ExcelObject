package com.str.newExcel.analysis.data;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

public interface StreamAnalysis {

    HSSFWorkbook steamToExcel(Object input);

    void excelToSteam(Object outPut, HSSFWorkbook excel) throws IOException;

}
