package com.cte.entity;

import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 描述:poi根据模板导出excel,根据excel坐标赋值,如(B1)
 */
public class ExcelExportUtil {

    //模板map
    private Map<String, Workbook> tempWorkbook = new HashMap<String, Workbook>();
    //模板输入流map
    private Map<String, InputStream> tempStream = new HashMap<String, InputStream>();

    /**
     * 功能:按模板向Excel中相应地方填充数据
     */
    public void writeData(String templateFilePath, Map<String, Object> dataMap, int sheetNo) throws IOException, InvalidFormatException {
        if (dataMap == null || dataMap.isEmpty()) {
            return;
        }
        //读取模板
        Workbook wbModule = getTempWorkbook(templateFilePath);
        //数据填充的sheet
        Sheet wsheet = wbModule.getSheetAt(sheetNo);

        for (Entry<String, Object> entry : dataMap.entrySet()) {
            String point = entry.getKey();
            Object data = entry.getValue();

            TempCell cell = getCell(point, data, wsheet);
            //指定坐标赋值
            setCell(cell, wsheet);
        }

        //设置生成excel中公式自动计算
        wsheet.setForceFormulaRecalculation(true);
    }

    /**
     * 功能:按模板向Excel中列表填充数据.只支持列合并
     */
    public void writeDateList(String templateFilePath, String[] heads, List<Map<Integer, Object>> datalist, int sheetNo) throws IOException, InvalidFormatException {
        if (heads == null || heads.length <= 0 || CollectionUtils.isEmpty(datalist)) {
            return;
        }
        //读取模板
        Workbook wbModule = getTempWorkbook(templateFilePath);
        //数据填充的sheet
        Sheet wsheet = wbModule.getSheetAt(sheetNo);

        //列表数据模板cell
        List<TempCell> tempCells = new ArrayList<TempCell>(heads.length);
        for (String point : heads) {
            TempCell tempCell = getCell(point, null, wsheet);
            //取得合并单元格位置 -1:表示不是合并单元格
            int pos = isMergedRegion(wsheet, tempCell.getRow(), tempCell.getColumn());
            if (pos > -1) {
                CellRangeAddress range = wsheet.getMergedRegion(pos);
                tempCell.setColumnSize(range.getLastColumn() - range.getFirstColumn());
            }
            tempCells.add(tempCell);
        }
        //赋值
        for (int i = 0; i < datalist.size(); i++) {//数据行
            Map<Integer, Object> dataMap = datalist.get(i);
            for (int j = 0; j < tempCells.size(); j++) {//列
                TempCell tempCell = tempCells.get(j);
                tempCell.setData(dataMap.get(j + 1));
                setCell(tempCell, wsheet);
                tempCell.setRow(tempCell.getRow() + 1);
            }
        }
    }

    /**
     * 功能:获取输入工作区
     */
    private Workbook getTempWorkbook(String templateFilePath) throws IOException, InvalidFormatException {
        if (!tempWorkbook.containsKey(templateFilePath)) {
            InputStream inputStream = getInputStream(templateFilePath);
            tempWorkbook.put(templateFilePath, WorkbookFactory.create(inputStream));
        }
        return tempWorkbook.get(templateFilePath);
    }

    /**
     * 功能:获得模板输入流
     */
    private InputStream getInputStream(String templateFilePath) throws FileNotFoundException {
        if (!tempStream.containsKey(templateFilePath)) {
            tempStream.put(templateFilePath, new FileInputStream((templateFilePath)));
        }
        return tempStream.get(templateFilePath);
    }

    /**
     * 功能:获取单元格数据,样式(根据坐标:B3)
     */
    private TempCell getCell(String point, Object data, Sheet sheet) {
        TempCell tempCell = new TempCell();

        //得到列  字母
        String lineStr = "";
        String reg = "[A-Z]+";
        Pattern p = Pattern.compile(reg);
        Matcher m = p.matcher(point);
        while (m.find()) {
            lineStr = m.group();
        }
        //将列字母转成列号 根据ascii转换
        char[] ch = lineStr.toCharArray();
        int column = 0;
        for (int i = 0; i < ch.length; i++) {
            char c = ch[i];
            int post = ch.length - i - 1;
            int r = (int) Math.pow(10, post);
            column = column + r * ((int) c - 65);
        }
        tempCell.setColumn(column);

        //得到行号
        reg = "[1-9]+";
        p = Pattern.compile(reg);
        m = p.matcher(point);
        while (m.find()) {
            tempCell.setRow((Integer.parseInt(m.group()) - 1));
        }

        //获取模板指定单元格样式,设置到tempCell(写列表数据的时候用)
        Row rowIn = sheet.getRow(tempCell.getRow());
        if (rowIn == null) {
            rowIn = sheet.createRow(tempCell.getRow());
        }
        Cell cellIn = rowIn.getCell(tempCell.getColumn());
        if (cellIn == null) {
            cellIn = rowIn.createCell(tempCell.getColumn());
        }
        tempCell.setCellStyle(cellIn.getCellStyle());
        tempCell.setData(data);
        return tempCell;
    }

    /**
     * 功能:给指定坐标单元格赋值
     */
    private void setCell(TempCell tempCell, Sheet sheet) {
        if (tempCell.getColumnSize() > -1) {
            CellRangeAddress rangeAddress = mergeRegion(sheet, tempCell.getRow(), tempCell.getRow(), tempCell.getColumn(), tempCell.getColumn() + tempCell.getColumnSize());
            setRegionStyle(tempCell.getCellStyle(), rangeAddress, sheet);
        }

        Row rowIn = sheet.getRow(tempCell.getRow());
        if (rowIn == null) {
            copyRows(tempCell.getRow() - 1, tempCell.getRow() - 1, tempCell.getRow(), sheet);//复制上一行
            rowIn = sheet.getRow(tempCell.getRow());
        }
        Cell cellIn = rowIn.getCell(tempCell.getColumn());
        if (cellIn == null) {
            cellIn = rowIn.createCell(tempCell.getColumn());
        }
        //根据data类型给cell赋值
        if (tempCell.getData() instanceof String) {
            cellIn.setCellValue((String) tempCell.getData());
        } else if (tempCell.getData() instanceof Integer) {
            cellIn.setCellValue((int) tempCell.getData());
        } else if (tempCell.getData() instanceof Double) {
            cellIn.setCellValue((double) tempCell.getData());
        } else {
            cellIn.setCellValue((String) tempCell.getData());
        }
        //样式
        if (tempCell.getCellStyle() != null && tempCell.getColumnSize() == -1) {
            cellIn.setCellStyle(tempCell.getCellStyle());
        }
    }

    /**
     * 功能:写到输出流并移除资源
     */
    public void writeAndClose(String templateFilePath, OutputStream os) throws IOException, InvalidFormatException {
        if (getTempWorkbook(templateFilePath) != null) {
            getTempWorkbook(templateFilePath).write(os);
            tempWorkbook.remove(templateFilePath);
        }
        if (getInputStream(templateFilePath) != null) {
            getInputStream(templateFilePath).close();
            tempStream.remove(templateFilePath);
        }
    }

    /**
     * 功能:判断指定的单元格是否是合并单元格
     */
    private Integer isMergedRegion(Sheet sheet, int row, int column) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return i;
                }
            }
        }
        return -1;
    }

    /**
     * 功能:合并单元格
     */
    private CellRangeAddress mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress rang = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(rang);
        return rang;
    }

    /**
     * 功能:设置合并单元格样式
     */
    private void setRegionStyle(CellStyle cs, CellRangeAddress region, Sheet sheet) {
        for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) row = sheet.createRow(i);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                    cell.setCellValue("");
                }
                cell.setCellStyle(cs);
            }
        }
    }

    /**
     * 功能:copy rows
     */
    private void copyRows(int startRow, int endRow, int pPosition, Sheet sheet) {
        int pStartRow = startRow - 1;
        int pEndRow = endRow - 1;
        int targetRowFrom;
        int targetRowTo;
        int columnCount;
        CellRangeAddress region = null;
        int i;
        int j;
        if (pStartRow == -1 || pEndRow == -1) {
            return;
        }
        // 拷贝合并的单元格
        for (i = 0; i < sheet.getNumMergedRegions(); i++) {
            region = sheet.getMergedRegion(i);
            if ((region.getFirstRow() >= pStartRow)
                    && (region.getLastRow() <= pEndRow)) {
                targetRowFrom = region.getFirstRow() - pStartRow + pPosition;
                targetRowTo = region.getLastRow() - pStartRow + pPosition;
                CellRangeAddress newRegion = region.copy();
                newRegion.setFirstRow(targetRowFrom);
                newRegion.setFirstColumn(region.getFirstColumn());
                newRegion.setLastRow(targetRowTo);
                newRegion.setLastColumn(region.getLastColumn());
                sheet.addMergedRegion(newRegion);
            }
        }
        // 设置列宽
        for (i = pStartRow; i <= pEndRow; i++) {
            Row sourceRow = sheet.getRow(i);
            columnCount = sourceRow.getLastCellNum();
            if (sourceRow != null) {
                Row newRow = sheet.createRow(pPosition - pStartRow + i);
                newRow.setHeight(sourceRow.getHeight());
                for (j = 0; j < columnCount; j++) {
                    Cell templateCell = sourceRow.getCell(j);
                    if (templateCell != null) {
                        Cell newCell = newRow.createCell(j);
                        copyCell(templateCell, newCell);
                    }
                }
            }
        }
    }

    /**
     * 功能:copy cell,不copy值
     */
    private void copyCell(Cell srcCell, Cell distCell) {
        distCell.setCellStyle(srcCell.getCellStyle());
        if (srcCell.getCellComment() != null) {
            distCell.setCellComment(srcCell.getCellComment());
        }
        int srcCellType = srcCell.getCellType();
        distCell.setCellType(srcCellType);
    }

    /**
     * 描述:临时单元格数据
     */
    class TempCell {
        private int row;
        private int column;
        private CellStyle cellStyle;
        private Object data;
        //用于列表合并,表示几列合并
        private int columnSize = -1;

        public int getColumn() {
            return column;
        }

        public void setColumn(int column) {
            this.column = column;
        }

        public int getRow() {
            return row;
        }

        public void setRow(int row) {
            this.row = row;
        }

        public CellStyle getCellStyle() {
            return cellStyle;
        }

        public void setCellStyle(CellStyle cellStyle) {
            this.cellStyle = cellStyle;
        }

        public Object getData() {
            return data;
        }

        public void setData(Object data) {
            this.data = data;
        }

        public int getColumnSize() {
            return columnSize;
        }

        public void setColumnSize(int columnSize) {
            this.columnSize = columnSize;
        }
    }

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
        String templateFilePath = "C:/load.xlsx";
        File file = new File("C:/aaaa/load2.xlsx");
        OutputStream os = new FileOutputStream(file);

        ExcelExportUtil excel = new ExcelExportUtil();
        Map<String, Object> dataMap = new HashMap<String, Object>();
        dataMap.put("B1", "导入模板");
        dataMap.put("B2", "统计时间：2017/01/01 ");

        excel.writeData(templateFilePath, dataMap, 0);

        List<Map<Integer, Object>> datalist = new ArrayList<Map<Integer, Object>>();
        Map<Integer, Object> data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口");
        data.put(6, "代码开发");
        data.put(7, "3.17");

        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);

        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "3/10/17");
        data.put(2, "18:50");
        data.put(3, "19:00");
        data.put(4, "李子鹏");
        data.put(5, "新增项目键值对接口,供任务计时调用新增项目键值对接口");
        data.put(6, "代码开发");
        data.put(7, "3.17");
        datalist.add(data);
        data = new HashMap<Integer, Object>();
        data.put(0, "1");
        data.put(1, "");
        data.put(2, "");
        data.put(3, "");
        data.put(4, "");
        data.put(5, "");
        data.put(6, "");
        data.put(7, "");
        datalist.add(data);

        String[] heads = new String[]{"B4", "C4", "D4", "E4", "F4", "G4", "H4"};
        excel.writeDateList(templateFilePath, heads, datalist, 0);

        //写到输出流并移除资源
        excel.writeAndClose(templateFilePath, os);

        os.flush();
        os.close();
    }

}