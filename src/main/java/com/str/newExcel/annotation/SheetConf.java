package com.str.newExcel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface SheetConf {

    /**
     * 表格名称
     */
    public String sheetName() default "";

    /**
     *是否拥有标题 如果为false 则生成的表只包含数据不包含标题
     */
    public boolean isHaveTitle() default true;

}
