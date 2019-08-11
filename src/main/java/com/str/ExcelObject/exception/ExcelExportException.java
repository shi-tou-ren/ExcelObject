package com.str.ExcelObject.exception;

import java.io.IOException;

/**
 * excel文件导出异常
 * @since 1.1.t
 * @author wangchenchen
 */
public class ExcelExportException extends IOException {

    public ExcelExportException() {
        super("Excel文件导出异常");
    }

    public ExcelExportException(String message) {
        super(message);
    }

}
