package com.ck.excel.suport;

import lombok.AllArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.function.Function;

/**
 * @author QianNianXiaoYao
 * @serial 2019/4/2
 */
@Slf4j
public class ExcelUtil {

    /**
     * 获取 指定  curCol 列的数据
     * <p>
     * stream 为null返回空集合
     *
     * @param is            stream
     * @param curCol        第几列 数据
     * @param skipFirstLine 是否跳过第一行
     * @return list
     * @throws IOException            e
     * @throws InvalidFormatException e
     * @author QianNianXiaoYao
     */
    public static List<String> getColumnForData(InputStream is, int curCol, boolean skipFirstLine) throws IOException, InvalidFormatException {
        if (is == null) {
            return Collections.emptyList();
        }
        Workbook workbook = WorkbookFactory.create(is);
        List<String> list = new ArrayList<>();
        Iterator<Row> rowIterator = workbook.getSheetAt(curCol).rowIterator();
        // 判断是否跳过
        if (skipFirstLine) {
            rowIterator.next();
        }
        rowIterator.forEachRemaining(s -> {
            Cell cell = s.getCell(0);
            // 判断是否为正常数据
            if (org.springframework.util.StringUtils.hasText(cell.toString())) {
                list.add(cell.toString());
            }
        });
        workbook.close();
        return list;
    }

    /**
     * 将list数据转换成指定的excel
     * 如果生成的数据总条数大于 sheetMaxNum (单个sheet页最大数量) 或者大于一百万行
     * 则自动进行分为多个sheet页进行保存
     * <p>
     * 如果数据总条数为100000条, 分为十个sheet页进行保存, 将进行异步操作(速度会快)
     *
     * @param sheetName   sheetName
     * @param objects     data
     * @param os          os
     * @param sheetMaxNum 每个sheet最大行数 小于等于0/null 则为默认值
     *                    默认值: 一百万行
     * @param sp          sp
     * @param <T>         T
     * @throws IOException e
     * @author QianNianXiaoYao
     */
    public static <T> void createExcel(String sheetName, List<T> objects, OutputStream os, Integer sheetMaxNum, Function<T, List<TitleAndVal>> sp) throws IOException {

        @AllArgsConstructor
        final class Attr {
            private Sheet sheet;
            private List<T> ls;
        }
        // 最大sheet页数量16384
        final int SHEET_MAX = 16384;
        // 每个sheet最大行数 1048756
        int max = 1000000;
        if (sheetMaxNum != null && sheetMaxNum > 0) {
            max = sheetMaxNum;
        }
        Workbook workbook = new SXSSFWorkbook(2500);
        List<List<T>> lists = ListUtil.splitList(objects, max);
        // 构建sheet页
        if (lists.size() > SHEET_MAX)
            throw new RuntimeException("超出最大Sheet页数量");
        String sn = sheetName;
        List<Attr> attrList = new ArrayList<>();
        for (int i = 0; i < lists.size(); i++) {
            workbook.createSheet(sn);
            attrList.add(new Attr(workbook.getSheet(sn), lists.get(i)));
            // 当多个sheet时将传进来的 sheetName 加 i
            sn = sheetName + (i + 1);
        }
        // 多个sheet时异步执行
        attrList.parallelStream().forEach(ls -> new ExcelUtil().cre(ls.sheet, ls.ls, sp));
        // 写出数据
        workbook.write(os);
        workbook.close();
        log.info("Excel创建完成");
    }

    /**
     * 下载文件时设置 response 响应头和下载文件名称
     * 适用于 .xlsx 文件
     *
     * @param response response
     * @param fileName fileName
     * @return response
     * @author QianNianXiaoYao
     */
    public static HttpServletResponse downloadSetHead(HttpServletResponse response, String fileName) {
        //强制浏览器下载文件
        response.setContentType("application/force-download");
        String fn = null;
        try {
            fn = new String(fileName.getBytes("gbk"), "iso8859-1");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        response.addHeader("Content-Disposition", "attachment;fileName=" + fn + ".xlsx");

        return response;
    }

    /**
     * 文件名转为base64
     *
     * @param response response
     * @param fileName fileName
     * @return response
     * @author QianNianXiaoYao
     */
    public static HttpServletResponse downloadSetBase64Head(HttpServletResponse response, String fileName) {
        //强制浏览器下载文件
        response.setContentType("application/force-download");
        final Base64.Encoder encoder = Base64.getEncoder();
        final String fn = encoder.encodeToString(fileName.getBytes(StandardCharsets.UTF_8));
        response.addHeader("Content-Disposition", "attachment;fileName=" + fn + ".xlsx");
        return response;
    }


    /**
     * @param sheet sheet Obj
     * @param ls    操作数据
     * @param sp    .
     * @param <T>   T
     * @author QianNianXiaoYao
     */
    private <T> void cre(Sheet sheet, List<T> ls, Function<T, List<TitleAndVal>> sp) {
        Iterator<T> iterator = ls.iterator();
        int iRow = 0;
        while (iterator.hasNext()) {
            List<TitleAndVal> apply = sp.apply(iterator.next());
            // 初始化表头
            if (iRow == 0) {
                sheet.createRow(iRow);
                // 冻结第一行
                sheet.createFreezePane(1, 1, 1, 1);
                for (int i = 0; i < apply.size(); i++) {
                    sheet.getRow(0).createCell(i).setCellValue(apply.get(i).k);
                }
                iRow++;
            }
            // 填充数据
            Row row = sheet.createRow(iRow);
            for (int i = 0; i < apply.size(); i++) {
                row.createCell(i)
                        //如果对象为null转为空字符
                        .setCellValue(apply.get(i).v == null ? "" : apply.get(i).v + "");
            }
            iRow++;
        }
    }

    @AllArgsConstructor
    public final static class TitleAndVal {
        // 表头名称
        private String k;
        // 对应表头的值
        private Object v;
    }

}
