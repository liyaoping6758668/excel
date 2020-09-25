package com.ping.excel.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * @Author：haung
 * @Date：2019-09-21 11:21
 * @Description：Excel导出工具类，依赖于ClassUtil工具类
 */
public final class ExcelExport2 {

    /**
     * 将传入的数据导出excel表并下载
     * @param response 返回的HttpServletResponse
     * @param importlist 要导出的对象的集合
     * @param attributeNames 含有每个对象属性在excel表中对应的标题字符串的数组（请按对象中属性排序调整字符串在数组中的位置）
     */
    public static void export(HttpServletResponse response, List<?> importlist, String[] attributeNames,String sheetTitle) {
        //获取数据集
        List<?> datalist = importlist;
        HSSFWorkbook workbook = getCommon(sheetTitle, attributeNames);
        HSSFSheet sheet = workbook.getSheet(sheetTitle);
        HSSFCellStyle cellStyle = getCellStyle(workbook);

        //获取对象属性
        Field[] fields = ClassUtil.getClassAttribute(importlist.get(0));
        //获取对象get方法
        List<Method> methodList = ClassUtil.getMethodGet(importlist.get(0));

        //创建普通行
        for (int i = 0;i<datalist.size();i++){
            //因为第一行已经用于创建标题行，故从第二行开始创建
            Row row = sheet.createRow(i+2);
            //如果是第一行就让其为标题行
            Object targetObj = datalist.get(i);
            for (int j = 0;j<fields.length;j++){
                //创建列
                Cell cell = row.createCell(j);
                cell.setCellType(CellType.STRING);
                //
                try {
                    Object value = methodList.get(j).invoke(targetObj, new Object[]{});
                    cell.setCellValue(transCellType(value));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                }
                cell.setCellStyle(cellStyle);
            }
        }
        response.setContentType("application/octet-stream");
        //默认Excel名称
        response.setHeader("Content-Disposition", "attachment;fileName="+"test.xls");

        try {
            response.flushBuffer();
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static String transCellType(Object value){
        String str = null;
        if (value instanceof Date){
            Date date = (Date) value;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            str = sdf.format(date);
        }else{
            str = String.valueOf(value);
            if (str == "null"){
                str = "";
            }
        }

        return str;
    }


    /**
     * 功能模板（标题及表头）
     */
    private static HSSFWorkbook getCommon(String sheetTitle, String[] fields) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetTitle);

        // 设置列宽度
        for (int i = 0; i < fields.length; i++) {
            sheet.setColumnWidth(i, 16 * 256);
        }

        // 通用样式
        HSSFCellStyle cellStyle = getCellStyle(workbook);

        // 标题样式
        HSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.cloneStyleFrom(cellStyle);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        HSSFFont titleFont = workbook.createFont();
        titleFont.setFontName("楷体");
        titleFont.setBold(true);
        titleFont.setFontHeight((short) 14);
        titleFont.setFontHeightInPoints((short)24);//设置字体大小

        titleStyle.setFont(titleFont);
        // 表头样式
        HSSFCellStyle thStyle = workbook.createCellStyle();
        thStyle.cloneStyleFrom(titleStyle);
        thStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        thStyle .setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        HSSFFont thFont = workbook.createFont();
        thFont.setFontName("楷体");
        thFont.setBold(titleFont.getBold());
        thFont.setColor(IndexedColors.WHITE.getIndex());
        thStyle.setFont(thFont);

        // 创建标题样式、表格表头
        HSSFRow titleRow = sheet.createRow(0);
        HSSFRow thsRow = sheet.createRow(1);
        for (int i = 0; i < fields.length; i++) {
            HSSFCell title = titleRow.createCell(i);
            title.setCellStyle(titleStyle);
            HSSFCell th = thsRow.createCell(i);
            th.setCellValue(fields[i]);
            th.setCellStyle(thStyle);
        }

        // 绘制标题
        titleRow.setHeight((short) (26 * 20));
        HSSFCell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(sheetTitle);
        titleCell.setCellStyle(titleStyle);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, fields.length - 1));
        return workbook;

    }

    /**
     * 获取通用样式
     */
    private static HSSFCellStyle getCellStyle(HSSFWorkbook workbook) {
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        HSSFFont font = workbook.createFont();
        font.setFontName("楷体");
        cellStyle.setFont(font);
        return cellStyle;
    }
}

