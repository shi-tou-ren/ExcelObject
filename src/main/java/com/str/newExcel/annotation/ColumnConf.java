package com.str.newExcel.annotation;

import javax.annotation.Resources;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ColumnConf {

    /**
     * 列标题名称
     */
    public String titleName() default "";


}
