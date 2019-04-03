package com.ck.excel.annotations;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author QianNianXiaoYao
 * @serial 2019/4/2
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface TitleAndVal {

    /**
     * 表头名称
     *
     * @return ""
     */
    String titleName();

    /**
     * 表名称映射实体类的字段名称
     *
     * @return ""
     */
    String mapField();

    /**
     * 例: "0=女|1=男"
     */
    String simpleExpression() default "";

//    Class<? extends ExcelExpression> clazz() default ExcelExpression.class;

}
