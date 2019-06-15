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

/**
 * Excel文件的对象
 * @see Sheet
 * @since 1.0.t
 * @author wangchenchen
 */
public class Excel {

    private ArrayList<Sheet> sheets = new ArrayList<Sheet>(1);

    private HSSFWorkbook workbook;

    /**
     * 创建空Excel文件对象
     */
    public Excel() {
        this.workbook = new HSSFWorkbook();
    }

    /**
     * 通过配置在此Excel中创建空白表 (未指定名称会随机生成名称)
     * @param config 被创建表的配置
     * @return 创建的表格的对象表示
     * @see Sheet
     */
    public Sheet createSheet(String config) {
        HSSFSheet hssfSheet = this.workbook.createSheet();
        Sheet sheet = new Sheet(hssfSheet, initSheetAndFormatConfig(hssfSheet, config));
        this.sheets.add(sheet);
        return sheet;
    }

    /**
     * 通过配置在此Excel中创建指定名称的空白表
     * @param name 被创建表的名称
     * @param config 被创建表的配置
     *               配置文件格式为：{id:id,name:名称,time:时间}或 id:id,name:名称,time:时间；其中键值对分别表示 成员属性名:表格列名称
     * @return 创建的表格的对象表示
     * @see Sheet
     * @throws  SheetCreateException 在方法运行过程中如果创建的表名称已经存在将会抛出此异常
     */
    public Sheet createSheet(String name, String config) {
        HSSFSheet hssfSheet = this.workbook.getSheet(name);
        if (null != hssfSheet) {
            throw new SheetCreateException("该名称的表已经存在");
        }
        hssfSheet = this.workbook.createSheet(name);
        Sheet sheet = new Sheet(hssfSheet, initSheetAndFormatConfig(hssfSheet, config));
        this.sheets.add(sheet);
        return sheet;
    }

    /**
     * 删除此Excel文件中指定下标的表格
     * @param index 指定表格的下标
     */
    public void deleteSheet(int index) {
        this.workbook.removeSheetAt(index);
        this.sheets.remove(index);
    }

    /**
     * 删除此Excel文件中指定名称的表格
     * @param sheetName 指定的表格名称
     */
    public void deleteSheet(String sheetName) {
        int index = this.workbook.getSheetIndex(sheetName);
        deleteSheet(index);
    }

    /**
     * 删除指定表格在此Excel文件中的存在
     * @param sheet 指定的表格
     */
    public void deleteSheet(Sheet sheet) {
        deleteSheet(sheet.getName());
    }

    /**
     * 通过下标获取相应的表格 如果不存在则返回null
     * @param index 下标
     * @return 指定下标下的表格
     * @see Sheet
     */
    public Sheet getSheet(int index) {
        if (index < 0 || index >= this.sheets.size()) {
            return null;
        }
        return this.sheets.get(index);
    }

    /**
     * 通过表名称获取相应的表格 如果不存在则返回null
     * @param sheetName 执行获取的表格名称
     * @return 指定名称的表格
     * @see Sheet
     */
    public Sheet getSheet(String sheetName) {
        return getSheet(this.workbook.getSheetIndex(sheetName));
    }

    /**
     * 导出Excel为指定文件 注意导出文件格式为.xls
     * @param file 指定的文件目录
     * @throws FileNotFoundException 目录异常
     */
    public void export(File file) throws FileNotFoundException {
        this.export(new FileOutputStream(file));
    }

    /**
     * 导出Excel到输出流中
     * @param outputStream 指定的输出流
     */
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


    /**
     * Excel表格的对象
     * 此类用于根据创建时的配置文件持续对Excel表格进行数据添加
     * @see Excel
     * @since 1.0.t
     * @author wangchenchen
     */
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
        /**
         * 获取当前表的名称
         * @return 当前表的名称
         */
        public String getName() {
            return this.hssfSheet.getSheetName();
        }

        /**
         * 对当前表进行重命名
         * @param sheetName 重命名之后的表名称
         * @throws SheetConfigException 在方法运行中如果新的表名已经被同一个Excel文件中的其它表使用将会抛出此异常
         */
        public void rename(String sheetName) {
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
         * 对该表格添加数据
         * @param data 添加的数据
         */
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
                                if (null != value) {
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
