package com.cte.entity;

import com.cte.controller.FileUploadController;
import com.cte.entity.util.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

/***
 *
 *
 * 描    述：导出文件为excel 2007版
 *
 * 创 建 者：@author wl
 * 创建时间：2018/12/515:43
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
public class ExportExcel2007Utils {

    private static final Logger logger = LoggerFactory.getLogger(FileUploadController.class);


    /*
        List<KeyValue> keyValues = new ArrayList<>();
        keyValues.add(new KeyValue("payUser", "付款方"));
        keyValues.add(new KeyValue("payId", "付款方ID"));
        keyValues.add(new KeyValue("receiptUser", "收款方"));
        keyValues.add(new KeyValue("receiptId", "收款方ID"));
    */

    /**
     * 2007excel导出
     *
     * @param fileName  文件名称
     * @param listMaps  类型的数据集(数据)
     * @param transDate 时间(默认为null)
     * @param keyValues 类型的表头
     * @param request
     * @param response
     */
    public void write2007(String fileName, String transDate, List<Map<String, Object>> listMaps, List<KeyValue> keyValues,
                          HttpServletRequest request, HttpServletResponse response) {

        //创建工作簿对象
        XSSFWorkbook wb = new XSSFWorkbook();
        OutputStream os = null;

        String downName = fileName;
        logger.info(downName + "下载开始");
        try {
            //为下载的文件名和Sheet编码设置为UTF-8
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            if (StringUtils.isBlank(transDate)) {
                transDate = new SimpleDateFormat("yyyy/MM/dd").format(new Date());
            }
            //区分IE浏览器和其他浏览器
            if (request.getHeader("User-Agent").contains("MSIE") || request.getHeader("User-Agent").contains("Trident")) {
                fileName = java.net.URLEncoder.encode((downName + transDate + ".xlsx"), "UTF-8");
            } else {
                fileName = new String((downName + transDate + ".xlsx").getBytes("UTF-8"), "ISO-8859-1");
            }
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName);
            //设置输出流对象
            os = response.getOutputStream();
            //根据入参title创建Sheet对象
            // set sheet
            XSSFSheet sheet = wb.createSheet(downName);
            //创建Sheet对象的第一行
            XSSFRow xssfRow = sheet.createRow(0);
            //根据入参keyValueList设置表头
            // set header title
            for (int i = 0; i < keyValues.size(); i++) {
                XSSFCell cell = xssfRow.createCell(i);
                cell.setCellValue(keyValues.get(i).getValue());
            }
            //从第二行开始按照表头创建行对象
            for (int i = 1; i <= listMaps.size(); i++) {
                Map<String, Object> map = listMaps.get(i - 1);
                XSSFRow rowValue = sheet.createRow(i);
                for (int j = 0; j < keyValues.size(); j++) {
                    KeyValue keyValue = keyValues.get(j);
                    XSSFCell cell = rowValue.createCell(j);
                    //通道业务类型
                    if (null != map.get(keyValue.getKey())) {
                        Object obj = map.get(keyValue.getKey());
                        if (obj instanceof Date) {
                            logger.info("");
                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy/MM/dd");
                            Date date = (Date) obj;
                            String formatTime = simpleDateFormat.format(date).replace("00:00:00.0", "").replace(".0", "");
                            cell.setCellValue(formatTime);
                        } else if (obj instanceof BigDecimal) {
                            DecimalFormat df = new DecimalFormat("#0.00");
                            cell.setCellValue(df.format(obj));
                        } else {
                            cell.setCellValue(obj.toString());
                        }
                    } else {
                        cell.setCellValue(" ");
                    }
                }
            }
            //写入输出流对象，导出
            wb.write(os);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                wb.close();
                os.close();
            } catch (IOException ie) {
                logger.error(fileName + "下载错误--->{}", ie);
            } catch (Exception e2) {
                logger.error(fileName + "下载错误--->{}", e2);
            }
        }
    }
}
