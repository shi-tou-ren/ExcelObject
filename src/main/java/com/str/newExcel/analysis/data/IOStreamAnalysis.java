package com.str.newExcel.analysis.data;

import com.str.newExcel.exception.AnalysisException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

public class IOStreamAnalysis implements StreamAnalysis {
    public HSSFWorkbook steamToExcel(Object input) {
        return null;
    }

    public void excelToSteam(Object outPut, HSSFWorkbook excel) throws IOException {
        if (!(outPut instanceof OutputStream)) {
            throw new AnalysisException("输出错误请使用OutPutStream");
        }
        excel.write((OutputStream) outPut);
        ((OutputStream) outPut).flush();
        ((OutputStream) outPut).close();
    }
}
