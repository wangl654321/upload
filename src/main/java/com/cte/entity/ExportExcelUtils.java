package com.cte.entity;

import com.cte.controller.FileUploadController;
import org.apache.poi.hssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
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
 * 描    述：文件导出Utils
 *
 * 创 建 者： @author wl
 * 创建时间： Nov 15, 201612:13:59 PM
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
public class ExportExcelUtils {

    private static final Logger logger = LoggerFactory.getLogger(FileUploadController.class);

	 /*
	 	使用方法
	 	// 表格生产的标题行
		String[] headers = {"付款方", "付款方ID", "收款方", "收款方ID", "转账日期", "转账订单号", "订单编号", "订单金额", "订单状态", "交易类型", "平台类型", "下单日期", "时间"};
		// 表格插入对应的字段值
		String[] showField = {"payUser", "payId", "receiptUser", "receiptId", "transDate", "transNo", "orderNo", "orderAmount", "orderStatus", "transType", "platformType", "date", "time"};
		excelController.exportExcel(source + "列表", headerArray, listExcel, showField, request, response);
	*/

    /**
     * @方法说明：文件导出Utils
     * @时间： 2018-01-29 15:03
     * @创建人：wanglObjToMap
     */
    public void exportExcel(String fileName, String[] headers, List<Map<String, Object>> dataSet,
                            String[] showField, HttpServletRequest request, HttpServletResponse response) {

        OutputStream out = null;
        //为下载的文件名和Sheet编码设置为UTF-8
        String downName = fileName;
        logger.info(downName + "下载开始");
        try {
            out = response.getOutputStream();
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            //区分IE浏览器和其他浏览器
            if (request.getHeader("User-Agent").contains("MSIE") || request.getHeader("User-Agent").contains("Trident")) {
                fileName = java.net.URLEncoder.encode((fileName + ".xlsx"), "UTF-8");
            } else {
                fileName = new String((fileName + ".xlsx").getBytes("UTF-8"), "ISO-8859-1");
            }
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName);

            // 声明一个工作薄
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 生成一个Sheet
            HSSFSheet sheet = workbook.createSheet(downName);

            //产生表格标题行
            HSSFRow row = sheet.createRow(0);
            for (short i = 0; i < headers.length; i++) {
                HSSFCell cell = row.createCell(i);
                HSSFRichTextString text = new HSSFRichTextString(headers[i]);
                cell.setCellValue(text);
            }

            int index = 0;
            for (Map<String, Object> valMap : dataSet) {
                index++;
                row = sheet.createRow(index);
                for (short i = 0; i < showField.length; i++) {
                    HSSFCell cell = row.createCell(i);
                    //通道业务类型
                    if (null != valMap.get(showField[i])) {
                        Object obj = valMap.get(showField[i]);
                        if (obj instanceof Date) {
                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy/MM/dd");
                            Date date = (Date) obj;
                            String format = simpleDateFormat.format(date).replace("00:00:00.0", "").replace(".0", "");
                            cell.setCellValue(format);
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
            workbook.write(out);
            logger.info(downName + "下载结束");
        } catch (IOException e) {
            logger.error(downName + "下载错误--->{}", e);
        } catch (Exception e) {
            logger.error(downName + "下载错误--->{}", e);
        } finally {
            try {
                out.flush();
                out.close();
            } catch (IOException e) {
                logger.error(downName + "下载错误--->{}", e);
            }
        }
    }
}
