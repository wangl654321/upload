package com.cte.entity;

/**
 * Created by User on 2017/12/13.
 */

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

//import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.springframework.stereotype.Component;

/**
 * 导出Excel公共方法
 *
 * @author wkr
 */

public class ExcelUtils {

    private HSSFWorkbook workbook = null;
    private FileOutputStream fileOut = null;
    private int sheeNum;

    //构造方法，传入要导出到文件
    public ExcelUtils(String fileName) throws FileNotFoundException {
        // 创建工作簿对象
        workbook = new HSSFWorkbook();
        fileOut = new FileOutputStream(fileName);
        sheeNum = 0;
    }

    public void excelFinal() throws IOException {
        workbook.write(fileOut);
        fileOut.close();
    }

    /*
     * 导出数据
     * */
    public void createSheet(String sheetName, String titleName, String[] queryName, String[] rowName, List<Object[]> dataList) throws Exception {
        try {
            // 创建工作表
            HSSFSheet sheet = workbook.createSheet();
            workbook.setSheetName(sheeNum, sheetName);
            sheeNum = sheeNum + 1;

            int m = 0;
            //设置标题行
            HSSFRow rowm = sheet.createRow(m);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(m, m, 0, rowName.length - 1);
            sheet.addMergedRegion(cellRangeAddress);
            HSSFCell cellTiltle = rowm.createCell(m);
            cellTiltle.setCellValue(titleName);
            //获取标题样式对象
            HSSFCellStyle titleStyle = this.getTitleStyle(workbook);
            cellTiltle.setCellStyle(titleStyle);
            this.setCellRangeAddressStyle(cellRangeAddress, sheet, workbook);
            m = m + 1;

            //设置查询统计行//获取查询样式对象
            HSSFCellStyle queryStyle = this.getQueryStyle(workbook);
            for (int i = 0; i < queryName.length - 1; i = i + 2) {
                HSSFRow queryRow = sheet.createRow(m);
                if (i == 0) {
                    queryRow.setHeight((short) 500);
                }
                CellRangeAddress cellRangeAddress1 = new CellRangeAddress(m, m, 0, 2);
                sheet.addMergedRegion(cellRangeAddress1);
                HSSFCell cellQueryColumnOne = queryRow.createCell(0);
                cellQueryColumnOne.setCellValue(queryName[i]);
                cellQueryColumnOne.setCellStyle(queryStyle);
                this.setCellRangeAddressStyle(cellRangeAddress1, sheet, workbook);
                CellRangeAddress cellRangeAddress2 = new CellRangeAddress(m, m, 3, rowName.length - 1);
                sheet.addMergedRegion(cellRangeAddress2);
                HSSFCell cellQueryColumnTwo = queryRow.createCell(3);
                cellQueryColumnTwo.setCellValue(queryName[i + 1]);
                cellQueryColumnTwo.setCellStyle(queryStyle);
                this.setCellRangeAddressStyle(cellRangeAddress2, sheet, workbook);
                m = m + 1;
            }

            if (queryName.length % 2 == 1) {
                HSSFRow queryRow = sheet.createRow(m);
                sheet.addMergedRegion(new CellRangeAddress(m, m, 0, rowName.length - 1));
                HSSFCell cellQueryColumnOne = queryRow.createCell(0);
                cellQueryColumnOne.setCellValue(queryName[queryName.length - 1]);
                cellQueryColumnOne.setCellStyle(queryStyle);
                m = m + 1;
            }

            //sheet样式定义【getColumnTopStyle()/getStyle()均为自定义方法 - 在下面  - 可扩展】
            //获取列头样式对象
            HSSFCellStyle columnTopStyle = this.getColumnTopStyle(workbook);
            //单元格样式对象
            HSSFCellStyle style = this.getStyle(workbook);

            /*
             *  产生表格标题行
            sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, (rowName.length-1))); //设置开头两行合并
            cellTiltle.setCellStyle(columnTopStyle);  //设置列头单元格样式
            cellTiltle.setCellValue("");//设置空单元格的内容。比如传入List的值是空的。
             */
            // 定义所需列数
            int columnNum = rowName.length;
            // 在索引2的位置创建行(最顶端的行开始的第二行)
            HSSFRow rowRowName = sheet.createRow(m);
            // 将列头设置到sheet的单元格中
            for (int n = 0; n < columnNum; n++) {
                //创建列头对应个数的单元格
                HSSFCell cellRowName = rowRowName.createCell(n);
                //设置列头单元格的数据类型
                cellRowName.setCellType(HSSFCell.CELL_TYPE_STRING);
                HSSFRichTextString text = new HSSFRichTextString(rowName[n]);
                //设置列头单元格的值
                cellRowName.setCellValue(text);
                //设置列头单元格样式
                cellRowName.setCellStyle(columnTopStyle);
            }

            m = m + 1;
            //将查询出的数据设置到sheet对应的单元格中
            for (int i = 0; i < dataList.size(); i++) {
                Object[] obj = dataList.get(i);
                //创建所需的行数
                HSSFRow row = sheet.createRow(i + m);
                for (int j = 0; j < obj.length; j++) {
                    //设置单元格的数据类型
                    HSSFCell cell = null;
                    cell = row.createCell(j, HSSFCell.CELL_TYPE_STRING);
                    if (!"".equals(obj[j]) && obj[j] != null) {
                        //设置单元格的值
                        cell.setCellValue(obj[j].toString());
                    } else {
                        //设置单元格的占位值
                        cell.setCellValue(" ");
                    }
                    //设置单元格样式
                    cell.setCellStyle(style);
                }
            }
            //让列宽随着导出的列长自动适应
            for (int colNum = 0; colNum < columnNum; colNum++) {
                sheet.autoSizeColumn(colNum, true);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void setCellRangeAddressStyle(CellRangeAddress cellRangeAddress, HSSFSheet sheet, HSSFWorkbook workbook) {
        RegionUtil.setBorderBottom(1, cellRangeAddress, sheet, workbook);
        RegionUtil.setBorderLeft(1, cellRangeAddress, sheet, workbook);
        RegionUtil.setBorderRight(1, cellRangeAddress, sheet, workbook);
        RegionUtil.setBorderTop(1, cellRangeAddress, sheet, workbook);
        RegionUtil.setBottomBorderColor(HSSFColor.WHITE.index, cellRangeAddress, sheet, workbook);
        RegionUtil.setLeftBorderColor(HSSFColor.WHITE.index, cellRangeAddress, sheet, workbook);
        RegionUtil.setRightBorderColor(HSSFColor.WHITE.index, cellRangeAddress, sheet, workbook);
        RegionUtil.setTopBorderColor(HSSFColor.WHITE.index, cellRangeAddress, sheet, workbook);
    }

    /*
     * 标题样式
     */
    public HSSFCellStyle getTitleStyle(HSSFWorkbook workbook) {
        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 11);
        //字体加粗
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字
        //font.setFontName("Courier New");
        font.setFontName("宋体");
        //设置样式;
        HSSFCellStyle style = workbook.createCellStyle();
        //设置底边框;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        //设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.WHITE.index);
        //设置左边框;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.WHITE.index);
        //设置右边框;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色;
        style.setRightBorderColor(HSSFColor.WHITE.index);
        //设置顶边框;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.WHITE.index);
        //在样式用应用设置的字体;
        style.setFont(font);
        //设置自动换行;
        style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }

    /*
     * 查询样式
     */
    public HSSFCellStyle getQueryStyle(HSSFWorkbook workbook) {
        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 12);
        //字体加粗
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字
        //font.setFontName("Courier New");
        font.setFontName("宋体");
        //设置样式;
        HSSFCellStyle style = workbook.createCellStyle();
        //设置底边框;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        //设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.WHITE.index);
        //设置左边框;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.WHITE.index);
        //设置右边框;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色;
        style.setRightBorderColor(HSSFColor.WHITE.index);
        //设置顶边框;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.WHITE.index);
        //在样式用应用设置的字体;
        style.setFont(font);
        //设置自动换行;
        style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        //设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }

    /*
     * 列头单元格样式
     */
    public HSSFCellStyle getColumnTopStyle(HSSFWorkbook workbook) {

        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 11);
        //字体加粗
        //font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字
        //font.setFontName("Courier New");
        font.setFontName("宋体");
        //设置样式;
        HSSFCellStyle style = workbook.createCellStyle();
        //设置底边框;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        //设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        //设置左边框;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        //设置右边框;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色;
        style.setRightBorderColor(HSSFColor.BLACK.index);
        //设置顶边框;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.BLACK.index);
        //在样式用应用设置的字体;
        style.setFont(font);
        //设置自动换行;
        style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }

    /*
     * 列数据信息单元格样式
     */
    public HSSFCellStyle getStyle(HSSFWorkbook workbook) {
        // 设置字体
        HSSFFont font = workbook.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 11);
        //字体加粗
        //font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        //设置字体名字
        //font.setFontName("Courier New");
        font.setFontName("宋体");
        //设置样式;
        HSSFCellStyle style = workbook.createCellStyle();
        //设置底边框;
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        //设置底边框颜色;
        style.setBottomBorderColor(HSSFColor.BLACK.index);
        //设置左边框;
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        //设置左边框颜色;
        style.setLeftBorderColor(HSSFColor.BLACK.index);
        //设置右边框;
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //设置右边框颜色;
        style.setRightBorderColor(HSSFColor.BLACK.index);
        //设置顶边框;
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        //设置顶边框颜色;
        style.setTopBorderColor(HSSFColor.BLACK.index);
        //在样式用应用设置的字体;
        style.setFont(font);
        //设置自动换行;
        style.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐;
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);

        return style;

    }
}
