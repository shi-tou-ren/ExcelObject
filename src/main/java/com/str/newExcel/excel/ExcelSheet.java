package com.str.newExcel.excel;

import com.str.newExcel.analysis.AnalysisFacade;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class ExcelSheet {

    private HSSFSheet sheet;

    private String[] columnMapping;

    ExcelSheet(HSSFSheet sheet, Object config) {
        this.sheet = sheet;
        columnMapping = AnalysisFacade.initSheet(sheet, config);
    }

    public ExcelSheet addData(Object data) throws IllegalAccessException {
        AnalysisFacade.dataToExcelSheet(sheet, columnMapping, data);
        return this;
    }
}
