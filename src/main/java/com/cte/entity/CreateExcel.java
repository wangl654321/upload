package com.cte.entity;

/**
 * @ClassName: CreateExcel
 * @Description: TODO()
 * @author www.xiongge.club
 * @date 2016-12-7 上午10:03:29
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

/**
 * @author Gerrard
 * @Discreption 根据已有的Excel模板，修改模板内容生成新Excel
 */
public class CreateExcel {

    /**
     * (2003 xls后缀 导出)
     */
    public static void createXLS() throws IOException {
        //excel模板路径
        File fi = new File("C:/load.xlsx");
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(fi));
        //读取excel模板
        HSSFWorkbook wb = new HSSFWorkbook(fs);
        //读取了模板内所有sheet内容
        HSSFSheet sheet = wb.getSheetAt(0);

        //如果这行没有了，整个公式都不会有自动计算的效果的
        sheet.setForceFormulaRecalculation(true);


        //在相应的单元格进行赋值
        HSSFCell cell = sheet.getRow(2).getCell(1);//第11行 第6列
        cell.setCellValue(1);
        HSSFCell cell2 = sheet.getRow(2).getCell(3);
        cell2.setCellValue(2);
        sheet.getRow(12).getCell(6).setCellValue(12);
        sheet.getRow(12).getCell(7).setCellValue(12);
        //修改模板内容导出新模板
        FileOutputStream out = new FileOutputStream("D:/export.xls");
        wb.write(out);
        out.close();
    }

    /**
     * (2007 xlsx后缀 导出)
     */
    public static void createXLSX() throws IOException {
        //excel模板路径
        File fi = new File("C:/load.xlsx");
        InputStream in = new FileInputStream(fi);
        //读取excel模板
        XSSFWorkbook wb = new XSSFWorkbook(in);
        //读取了模板内所有sheet内容
        XSSFSheet sheet = wb.getSheetAt(1);

        //如果这行没有了，整个公式都不会有自动计算的效果的
        sheet.setForceFormulaRecalculation(true);

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String format = simpleDateFormat.format(new Date()).replace("00:00:00.0", "").replace(".0", "");

        //在相应的单元格进行赋值
        XSSFCell cell = sheet.getRow(1).getCell(1);//第11行 第6列
        cell.setCellValue(format);
        XSSFCell cell2 = sheet.getRow(1).getCell(3);
        cell2.setCellValue("是");
        XSSFCell cell3 = sheet.getRow(1).getCell(5);
        cell3.setCellValue("建春");

        sheet.getRow(3).getCell(0).setCellValue("wang");
        sheet.getRow(3).getCell(1).setCellValue("lu");


        //修改模板内容导出新模板
        FileOutputStream out = new FileOutputStream("C:/aaaa/load2.xlsx");
        wb.write(out);
        out.close();
    }

    public static void main(String[] args) throws IOException {
        //excle 2003
        //createXLS();
        //excle 2007
        createXLSX();
    }

    public static XSSFCellStyle getBodyStyle(XSSFWorkbook wb) {
        // 创建单元格样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        // 设置单元格的背景颜色为淡蓝色
         cellStyle.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
         cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        // 设置单元格居中对齐
        // 设置单元格内容水平对其方式
        // XSSFCellStyle.ALIGN_CENTER 居中对齐
        // XSSFCellStyle.ALIGN_LEFT 左对齐
        // XSSFCellStyle.ALIGN_RIGHT 右对齐 cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        // 设置单元格垂直居中对齐
        // 设置单元格内容垂直对其方式
        // XSSFCellStyle.VERTICAL_TOP 上对齐
        // XSSFCellStyle.VERTICAL_CENTER 中对齐
        // XSSFCellStyle.VERTICAL_BOTTOM 下对齐 cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        // 创建单元格内容显示不下时自动换行
        cellStyle.setWrapText(true);
        // 设置单元格字体样式
        XSSFFont font = wb.createFont();
        // 设置字体加粗
        font.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL);
        font.setFontName("宋体");
        font.setFontHeight((short) 200);
        cellStyle.setFont(font);
        // 设置单元格边框为细线条
        // 设置单元格边框样式
        // CellStyle.BORDER_DOUBLE 双边线
        // CellStyle.BORDER_THIN 细边线
        // CellStyle.BORDER_MEDIUM 中等边线
        // CellStyle.BORDER_DASHED 虚线边线
        // CellStyle.BORDER_HAIR 小圆点虚线边线
        // CellStyle.BORDER_THICK 粗边线
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN); cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN); cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN); cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
        return cellStyle;
    }
}