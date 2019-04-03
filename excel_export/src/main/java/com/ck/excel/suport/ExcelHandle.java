package com.ck.excel.suport;

import com.ck.excel.annotations.CreateExcel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;
import org.springframework.stereotype.Component;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author QianNianXiaoYao
 * @serial 2019/4/2
 */
@Aspect
@Component
@Slf4j
public class ExcelHandle {

    private final static String splitStr = "[|]";
    private final static String parallelism = "=";

    private HttpServletResponse response;
    private List<TileMapField> tileMapFields;
    private String sheetName;
    private int sheetMaxNum;

    /**
     * 解析表达式
     *
     * @param str val
     * @return map
     * @author QianNianXiaoYao
     */
    private static Map<String, String> analyticalExpression(String str) {
        if (StringUtils.hasText(str)) {
            Map<String, String> map = new HashMap<>();
            String[] split = str.split(splitStr);
            Stream.of(split).forEach(it -> {
                String[] strings = it.split(parallelism);
                map.put(strings[0], strings[1]);
            });
            return map;
        }
        return null;
    }

    @Around(value = "@annotation(com.ck.excel.annotations.CreateExcel)")
    public Object around(ProceedingJoinPoint point) throws Throwable {
        Class<?> aClass = point.getTarget().getClass();
        String methodName = point.getSignature().getName();

        Stream.of(aClass.getMethods()).filter(it -> methodName.equals(it.getName())).findAny().ifPresent(it -> {
            CreateExcel annotation = it.getAnnotation(CreateExcel.class);
            _init(annotation);
        });

        Object proceed = point.proceed(point.getArgs());
        if (proceed instanceof List) {
            // noinspection unchecked
            generate((List<Object>) proceed);
        }
        return null;
    }

    /**
     * 初始化信息
     *
     * @param excel annotation
     * @author QianNianXiaoYao
     */
    private void _init(CreateExcel excel) {
        String fileName = excel.fileName();
        HttpServletResponse response = ((ServletRequestAttributes) (RequestContextHolder.currentRequestAttributes())).getResponse();
        this.response = excel.generatingCode().getCode(response, fileName);
        this.sheetName = excel.sheetName().equals("") ? fileName : excel.sheetName();
        this.sheetMaxNum = excel.sheetMaxNum();
        this.tileMapFields = Stream.of(excel.titleAndVal()).map(it -> new TileMapField(it.titleName(), it.mapField(), analyticalExpression(it.simpleExpression()))).collect(Collectors.toList());

    }

    /**
     * 清除头信息, 防止以文件方式下载错误信息
     *
     * @param response resp
     * @author QianNianXiaoYao
     */
    private void clearResponseHear(HttpServletResponse response) {
        response.setContentType("");
        response.setHeader("Content-Disposition", "");
    }

    /**
     * 生成 excel
     *
     * @param list val
     * @author QianNianXiaoYao
     */
    private void generate(List<Object> list) {
        // 没有找到数据抛出指定异常
        if (CollectionUtils.isEmpty(list)) {
            clearResponseHear(response);
            throw new RuntimeException("生成Excel数据没有找到!");
        }
        try {
            ExcelUtil.createExcel(sheetName, list, response.getOutputStream(), sheetMaxNum, this::getTitleAndValList);
        } catch (IOException e) {
            clearResponseHear(response);
            throw new RuntimeException("生成Excel失败!");
        }
    }

    /**
     * 生成TitleAndVal集合
     *
     * @param obj 目标对象
     * @return list
     * @author QianNianXiaoYao
     */
    private List<ExcelUtil.TitleAndVal> getTitleAndValList(Object obj) {
        List<ExcelUtil.TitleAndVal> tav = new ArrayList<>();
        Class<?> aClass = obj.getClass();
        for (TileMapField field : this.tileMapFields) {

            Object fieldVal = getFieldVal(field.getMapField(), obj, aClass);
            if (field.mapping == null) {
                tav.add(new ExcelUtil.TitleAndVal(field.titleName, fieldVal));
            } else {
                String s = field.mapping.get(fieldVal.toString());
                tav.add(new ExcelUtil.TitleAndVal(field.titleName, s));
            }
        }
        return tav;
    }

    /**
     * 获取指定字段值
     *
     * @param fieldName 字段名
     * @param _this     obj
     * @return val
     * @author QianNianXiaoYao
     */
    private Object getFieldVal(String fieldName, Object _this, Class<?> aClass) {
        Object val = null;
        try {
            Field field = aClass.getDeclaredField(fieldName);
            field.setAccessible(true);
            val = field.get(_this);
            // 加回访问权限
            field.setAccessible(false);
        } catch (NoSuchFieldException | IllegalAccessException e) {
            log.error("映射字段名不存在. " + fieldName, e);
        }
        return val;
    }

    /**
     * 表头名称和字段的映射
     *
     * @author QianNianXiaoYao
     */
    @Getter
    @AllArgsConstructor
    private class TileMapField {
        private String titleName;
        private String mapField;
        private Map<String, String> mapping;
    }

}









