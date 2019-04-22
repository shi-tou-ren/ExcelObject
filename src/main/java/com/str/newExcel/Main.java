package com.str.newExcel;

import com.str.newExcel.excel.Excel;
import com.str.newExcel.excel.ExcelSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by hasee on 2019/4/22.
 */
public class Main {

    public static void main(String[] args) throws IOException, IllegalAccessException {
        System.out.println(TestPojs.class.getSimpleName());
        TestPojs testPojs = new TestPojs();
        List<TestPojs> testPojs1 = new ArrayList<TestPojs>();
        testPojs1.add(testPojs);
        Excel newExvel = new Excel();
        ExcelSheet sheet = newExvel.createSheet(testPojs.getClass());
        sheet.addData(testPojs1);
        sheet.addData(testPojs1);


//        File file = new File("E:\\test.xls");
//        FileOutputStream fileOutputStream = new FileOutputStream(file);
        newExvel.outPut(new FileOutputStream(new File("E:\\test.xls")));
        newExvel.outPut(new FileOutputStream(new File("E:\\test1.xls")));

    }

}
