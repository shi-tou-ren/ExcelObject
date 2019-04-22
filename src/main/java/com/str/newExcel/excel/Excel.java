package com.str.newExcel.excel;

import com.str.newExcel.analysis.AnalysisFacade;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;
import java.util.List;

public class Excel {

    private List<ExcelSheet> sheets;

    private HSSFWorkbook workbook;

    public Excel() {
        workbook = new HSSFWorkbook();
    }

    public Excel(Object data) {
        workbook = AnalysisFacade.dataToExcel(data);
    }

    public ExcelSheet createSheet(Object config) {
        return new ExcelSheet(workbook.createSheet(), config);
    }

    public void outPut(Object outPut) throws IOException {
        AnalysisFacade.ExcelToStream(outPut, workbook);
    }

}
