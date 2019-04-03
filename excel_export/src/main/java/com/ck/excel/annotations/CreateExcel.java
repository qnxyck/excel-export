package com.ck.excel.annotations;


import com.ck.excel.suport.ExcelUtil;

import javax.servlet.http.HttpServletResponse;
import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 快速生成Excel
 * 要求方法返回值为List<T>集合
 *
 * @author QianNianXiaoYao
 * @serial 2019/4/2
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface CreateExcel {

    /**
     * 生成文件名称
     *
     * @return ""
     */
    String fileName();

    /**
     * sheet 名称
     * 如果不指定默认使用fileName
     *
     * @return ""
     */
    String sheetName() default "";

    /**
     * sheet页分割值 -1为sheet页最大值
     *
     * @return int
     */
    int sheetMaxNum() default -1;

    /**
     * 文件名生成编码
     *
     * @return code
     */
    GeneratingCode generatingCode() default GeneratingCode.DEFAULT;

    /**
     * 表头名称的指定
     *
     * @return titleAndVal
     */
    TitleAndVal[] titleAndVal() default {};

    enum GeneratingCode {
        BASE64() {
            @Override
            public HttpServletResponse getCode(HttpServletResponse response, String str) {
                return ExcelUtil.downloadSetBase64Head(response, str);
            }
        },
        DEFAULT() {
            @Override
            public HttpServletResponse getCode(HttpServletResponse response, String str) {
                return ExcelUtil.downloadSetHead(response, str);
            }
        };

        public abstract HttpServletResponse getCode(HttpServletResponse response, String str);
    }

}
