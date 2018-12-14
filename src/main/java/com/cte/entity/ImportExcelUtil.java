package com.cte.entity;

import com.cte.controller.FileUploadController;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/***
 *
 *
 * 描    述：java Excel导入、自适应版本、将Excel转成List<map>对象
 *
 * 创 建 者： @author wl
 * 创建时间： 2018-01-18 10:29
 * 创建描述：
 *
 * 修 改 者：  
 * 修改时间： 
 * 修改描述： 
 *
 * 审 核 者：
 * 审核时间：
 * 审核描述：
 *
 */
@Component
public class ImportExcelUtil {

    private static final Logger logger = LoggerFactory.getLogger(FileUploadController.class);

    /**
     * 2003- 版本的excel
     */
    private final static String XLS = ".xls";

    /**
     * 2007+ 版本的excel
     */
    private final static String X_LSX = ".xlsx";



     /*
        使用模板
        Map<String, String> map = new HashMap<>(16);
        map.put("订单编号","orderNo");
        map.put("订单金额","orderAmount");
        map.put("订单状态","orderStatus");

        map.put("交易类型","transType");
        map.put("平台类型","platformType");
    */

    /**
     * 将流中的Excel数据转成List<Map>
     *
     * @return
     * @throws Exception
     * @param-InputStream 输入流
     * @param-fileName 文件名（判断Excel版本）
     * @param-mapping 字段名称映射
     */
    public static List<Map<String, Object>> parseExcel(MultipartFile file, Map<String, String> map) throws Exception {

        logger.info(file.getName() + "文件导入开始");
        InputStream in = file.getInputStream();
        String fileName = file.getOriginalFilename();
        List<Map<String, Object>> list = null;
        try {
            // 根据文件名来创建Excel工作薄
            Workbook work = getWorkbook(in, fileName);
            if (null == work) {
                throw new Exception("创建Excel工作薄为空！");
            }
            Sheet sheet = null;
            Row row = null;
            Cell cell = null;
            // 返回数据
            list = new ArrayList<>();
            // 遍历Excel中所有的sheet
            for (int i = 0; i < work.getNumberOfSheets(); i++) {
                sheet = work.getSheetAt(i);
                if (sheet == null) {
                    continue;
                }
                // 取第一行标题
                row = sheet.getRow(0);
                String title[] = null;
                if (row != null) {
                    title = new String[row.getLastCellNum()];
                    for (int fast = row.getFirstCellNum(), last = row.getLastCellNum(); fast < last; fast++) {
                        cell = row.getCell(fast);
                        title[fast] = (String) getCellValue(cell);
                    }
                } else {
                    continue;
                }

                // 遍历当前sheet中的所有行
                for (int j = 1; j < sheet.getLastRowNum() + 1; j++) {
                    row = sheet.getRow(j);
                    Map<String, Object> m = new HashMap<>(row.getLastCellNum());
                    // 遍历所有的列
                    for (int fast = row.getFirstCellNum(), last = row.getLastCellNum(); fast < last; fast++) {
                        cell = row.getCell(fast);
                        String key = title[fast];
                        m.put(map.get(key), getCellValue(cell));
                    }
                    list.add(m);
                }
            }
        } catch (Exception e) {
            logger.info("java Excel导入失败,{}", e.getStackTrace());
        } finally {
            in.close();
        }
        return list;
    }

    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     *
     * @param inStr ,fileName
     * @return
     * @throws Exception
     */
    public static Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {

        Workbook workbook = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (XLS.equals(fileType)) {
            // 2003-
            workbook = new HSSFWorkbook(inStr);
        } else if (X_LSX.equals(fileType)) {
            // 2007+
            workbook = new XSSFWorkbook(inStr);
        } else {
            throw new Exception("解析的文件格式有误！");
        }
        return workbook;
    }

    /**
     * 描述：对表格中数值进行格式化
     *
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell) {

        Object value = null;
        // 格式化number String字符
        DecimalFormat df = new DecimalFormat("##0.00");
        // 日期格式化
        SimpleDateFormat sdf = new SimpleDateFormat("yyy/MM/dd");
        // 格式化数字
        DecimalFormat df2 = new DecimalFormat("##0.00");

        if (null == cell) {
            return "";
        }

        String General = "General";
        String timeStyle = "m/d/yy";
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (General.equals(cell.getCellStyle().getDataFormatString())) {
                    value = df.format(cell.getNumericCellValue());
                } else if (timeStyle.equals(cell.getCellStyle().getDataFormatString())) {
                    value = sdf.format(cell.getDateCellValue());
                } else {
                    value = df2.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                break;
        }
        return value;
    }

}
