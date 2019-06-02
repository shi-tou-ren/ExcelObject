package com.str.ExcelObject.exception;


public class SheetCreateException extends RuntimeException {

    public SheetCreateException() {
        super("Excel表创建异常");
    }

    public SheetCreateException(String message) {
        super(message);
    }

}
