package com.str.ExcelObject.excel;

import com.str.ExcelObject.exception.SheetConfigException;
import com.str.ExcelObject.exception.SheetCreateException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

public class Excel {

    private ArrayList<Sheet> sheets = new ArrayList<Sheet>(1);

    private HSSFWorkbook workbook;

    /**
     * -----------------------创建Excel文件-----------------------
     */
    public Excel() {
        this.workbook = new HSSFWorkbook();
    }

    /**
     * -----------------------创建Excel表格-----------------------
     */
    //创建空白表
    public Sheet createSheet(String config) {
        HSSFSheet hssfSheet = this.workbook.createSheet();
        Sheet sheet = new Sheet(hssfSheet, initSheetAndFormatConfig(hssfSheet, config));
        this.sheets.add(sheet);
        return sheet;
    }

    //创建指定名称的表
    public Sheet createSheet(String name, String config) {
        HSSFSheet hssfSheet = this.workbook.getSheet(name);
        if (null != hssfSheet) {
            //抛出错误该名称的表已经存在
            throw new SheetCreateException("该名称的表已经存在");
        }
        hssfSheet = this.workbook.createSheet(name);
        Sheet sheet = new Sheet(hssfSheet, initSheetAndFormatConfig(hssfSheet, config));
        this.sheets.add(sheet);
        return sheet;
    }

    /**
     * -----------------------删除Excel表格-----------------------
     */

    //根据索引删除表
    public void deleteSheet(int index) {
        this.workbook.removeSheetAt(index);
        this.sheets.remove(index);
    }

    //根据表名称删除表
    public void deleteSheet(String sheetName) {
        int index = this.workbook.getSheetIndex(sheetName);
        deleteSheet(index);
    }

    //删除指定的表本身
    public void deleteSheet(Sheet sheet) {
        deleteSheet(sheet.getName());
    }

    /**
     * -----------------------查找获取Excel表格-----------------------
     */
    //通过下标获取响应的表格 如果不存在则返回null
    public Sheet getSheet(int index) {
        if (index < 0 || index >= this.sheets.size()) {
            return null;
        }
        return this.sheets.get(index);
    }

    //通过表名称获取响应的表格 如果不存在则返回null
    public Sheet getSheet(String sheetName) {
        return getSheet(this.workbook.getSheetIndex(sheetName));
    }

    /**
     * -----------------------导出-----------------------
     */
    //导出为指定文件
    public void export(File file) throws FileNotFoundException {
        this.export(new FileOutputStream(file));
    }

    //导出到输出流中
    public void export(OutputStream outputStream) {
        try {
            this.workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String[] initSheetAndFormatConfig(HSSFSheet hssfSheet, String config) {
        if (hssfSheet.getLastRowNum() != 0) {
            hssfSheet.shiftRows(0, hssfSheet.getLastRowNum(), 1);
        }
        HSSFRow hssfRow = hssfSheet.createRow(0);
        String[] columnAndTitles = config.replaceAll("\\s|\\{|\\}", "").split(",");
        String[] columnMapping = new String[columnAndTitles.length];
        for (int i = 0; i < columnAndTitles.length; i++) {
            String[] cs = columnAndTitles[i].split(":");
            columnMapping[i] = cs[0];
            hssfRow.createCell(i).setCellValue(cs[1]);
        }
        return columnMapping;
    }


    public class Sheet {

        //列映射  关系为 下标->列名   LinkedHashSet
        private String[] columnMapping = null;

        private SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

        private HSSFSheet hssfSheet;

        Sheet(HSSFSheet hssfSheet, String[] columnMapping) {
            this.hssfSheet = hssfSheet;
            this.columnMapping = columnMapping;
        }

        /**
         * -----------------------表名称操作-----------------------
         */
        //获取表名称
        public String getName() {
            return this.hssfSheet.getSheetName();
        }

        //表重命名
        public void rename(String sheetName) {
            //如果重命名与原名相同则不操作
            if (!this.hssfSheet.getSheetName().equals(sheetName)) {
                HSSFWorkbook hssfWorkbook = this.hssfSheet.getWorkbook();
                HSSFSheet hssfSheet = hssfWorkbook.getSheet(sheetName);
                if (null != hssfSheet) {
                    throw new SheetConfigException("表重命名失败->表名已被使用");
                }
                hssfWorkbook.setSheetName(hssfWorkbook.getSheetIndex(this.hssfSheet), sheetName);
            }
        }

        /**
         *  -----------------------表样式操作-----------------------
         */

        /**
         *  -----------------------表初始化操作-----------------------
         */

        /**
         * -----------------------表数据操作-----------------------
         */
        //添加数据
        public void addData(Collection data) {

            //判断是否有数据
            if (null == data || data.size() == 0) {
                return;
            }

            if (data instanceof List) {
                addData(((List) data).get(0).getClass(), data);
            } else {
                for (Object o : data) {
                    addData(o.getClass(), data);
                    break;
                }
            }
        }

        private void addData (Class clazz, Collection data) {
            Field[] fields = new Field[columnMapping.length];
            for (int i = 0; i < fields.length; i++) {
                try {
                    fields[i] = clazz.getDeclaredField(columnMapping[i]);
                    fields[i].setAccessible(true);
                } catch (NoSuchFieldException e) {
                    //如果没有就跳过
                }
            }
            //getLastRowNum 获取的是最后一个数据的下标 所以开头是getLastRowNum的返回值加1
            int i = this.hssfSheet.getLastRowNum() + 1;
            for (Object o : data) {
                HSSFRow hssfRow = this.hssfSheet.createRow(i++);
                int cellIndex = 0;
                for (Field field : fields) {
                    try {
                        if (null != field) {
                            HSSFCell hssfCell = hssfRow.createCell(cellIndex++);
                            Object value = field.get(o);
                            if (!field.getType().isArray()) {
                                if (value instanceof String) {
                                    hssfCell.setCellValue((String) value);
                                } else if (value instanceof Number) {
                                    hssfCell.setCellValue(value.toString());
                                } else if (value instanceof Date) {
                                    hssfCell.setCellValue(this.dateFormat.format(value));
                                } else if (value instanceof Boolean) {
                                    hssfCell.setCellValue((Boolean) value);
                                } else {
                                    hssfCell.setCellValue(value.toString());
                                }
                            } else {
                                hssfCell.setCellValue(Arrays.toString((Object[]) value));
                            }
                        } else {
                            cellIndex++;
                        }
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                }
            }
        }

    }

}
