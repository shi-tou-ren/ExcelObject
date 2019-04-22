package com.str.newExcel;

import com.str.newExcel.annotation.ColumnConf;
import com.str.newExcel.annotation.SheetConf;

import java.util.Date;

@SheetConf
public class TestPojs {

    @ColumnConf(titleName = "测试")
    private int i = 10;

    private int j = 5;

    private String test = "测试文本";

    private Date date = new Date();

}
