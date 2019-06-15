package com.str.ExcelObject.exception;

/**
 * 表创建异常
 * @since 1.0.t
 * @author wangchenchen
 */
public class SheetCreateException extends RuntimeException {

    public SheetCreateException() {
        super("Excel表创建异常");
    }

    public SheetCreateException(String message) {
        super(message);
    }

}
