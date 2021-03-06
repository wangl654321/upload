package com.cte.entity;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.RichTextString;
/**
 * 共分为六部完成根据模板导出excel操作：<br/>
 * 第一步、设置excel模板路径（setSrcPath）<br/>
 * 第二步、设置要生成excel文件路径（setDesPath）<br/>
 * 第三步、设置模板中哪个Sheet列（setSheetName）<br/>
 * 第四步、获取所读取excel模板的对象（getSheet）<br/>
 * 第五步、设置数据（分为6种类型数据：setCellStrValue、setCellDateValue、setCellDoubleValue、setCellBoolValue、setCellCalendarValue、setCellRichTextStrValue）<br/>
 * 第六步、完成导出 （exportToNewFile）<br/>
 *
 * @author Administrator
 *
 */
public class ExcelWriter {
    POIFSFileSystem fs = null;
    HSSFWorkbook wb = null;
    HSSFSheet sheet = null;
    HSSFCellStyle cellStyle = null;

    private String srcXlsPath = "C:/load.xlsx";//  excel模板路径
    private String desXlsPath = "C:/aaaa/load2.xlsx";  // 生成路径
    private String sheetName = "load22";

    /**
     * 第一步、设置excel模板路径
     * @param srcXlsPaths
     */
    public void setSrcPath(String srcXlsPaths) {
        this.srcXlsPath = srcXlsPaths;
    }

    /**
     * 第二步、设置要生成excel文件路径
     * @param desXlsPaths
     * @throws FileNotFoundException
     */
    public void setDesPath(String desXlsPaths) throws FileNotFoundException {
        this.desXlsPath = desXlsPaths;
    }

    /**
     * 第三步、设置模板中哪个Sheet列
     * @param sheetName
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * 第四步、获取所读取excel模板的对象
     */
    public void getSheet() {
        try {
            File fi = new File(srcXlsPath);
            if(!fi.exists()){
                //System.out.println("模板文件:"+srcXlsPath+"不存在!");
                return;
            }
            fs = new POIFSFileSystem(new FileInputStream(fi));
            wb = new HSSFWorkbook(fs);
            sheet = wb.getSheet(sheetName);

            //生成单元格样式
            cellStyle = wb.createCellStyle();
            //设置背景颜色
            cellStyle.setFillForegroundColor(HSSFColor.RED.index);
            //solid 填充  foreground  前景色
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    /**
     *
     */
    public HSSFRow createRow(int rowIndex) {
        HSSFRow row = sheet.createRow(rowIndex);
        return row;
    }
    /**
     *
     */
    public void createCell(HSSFRow row,int colIndex) {
        row.createCell(colIndex);
    }
    /**
     * 第五步、设置单元格的样式
     * @param rowIndex    行值
     * @param cellnum    列值
     */
    public void setCellStyle(int rowIndex, int cellnum) {
        HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 第五步、设置字符串类型的数据
     * @param rowIndex    行值
     * @param cellnum    列值
     * @param value        字符串类型的数据
     */
    public void setCellStrValue(int rowIndex, int cellnum, String value) {
        if(value != null) {
            HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
            cell.setCellValue(value);
        }
    }

    /**
     * 第五步、设置日期/时间类型的数据
     * @param rowIndex    行值
     * @param cellnum    列值
     * @param value        日期/时间类型的数据
     */
    public void setCellDateValue(int rowIndex, int cellnum, Date value) {
        HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
        cell.setCellValue(value);
    }

    /**
     * 第五步、设置浮点类型的数据
     * @param rowIndex    行值
     * @param cellnum    列值
     * @param value        浮点类型的数据
     */
    public void setCellDoubleValue(int rowIndex, int cellnum, double value) {
        HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
        cell.setCellValue(value);
    }

    /**
     * 第五步、设置Bool类型的数据
     * @param rowIndex    行值
     * @param cellnum    列值
     * @param value        Bool类型的数据
     */
    public void setCellBoolValue(int rowIndex, int cellnum, boolean value) {
        HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
        cell.setCellValue(value);
    }

    /**
     * 第五步、设置日历类型的数据
     * @param rowIndex    行值
     * @param cellnum    列值
     * @param value        日历类型的数据
     */
    public void setCellCalendarValue(int rowIndex, int cellnum, Calendar value) {
        HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
        cell.setCellValue(value);
    }

    /**
     * 第五步、设置富文本字符串类型的数据。可以为同一个单元格内的字符串的不同部分设置不同的字体、颜色、下划线
     * @param rowIndex    行值
     * @param cellnum    列值
     * @param value        富文本字符串类型的数据
     */
    public void setCellRichTextStrValue(int rowIndex, int cellnum,
                                        RichTextString value) {
        HSSFCell cell = sheet.getRow(rowIndex).getCell(cellnum);
        cell.setCellValue(value);
    }

    /**
     * 第六步、完成导出
     */
    public void exportToNewFile() {
        FileOutputStream out;
        try {
            out = new FileOutputStream(desXlsPath);
            wb.write(out);
            out.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args){
        ExcelWriter excelWriter = new ExcelWriter();
        excelWriter.getSheet();
    }

}