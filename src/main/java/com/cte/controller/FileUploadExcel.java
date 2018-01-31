package com.cte.controller;

import com.cte.entity.ImportExcelUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/***
 *
 *
 * 描    述：文件上传
 *
 * 创 建 者： @author wangl
 * 创建时间： 2017-11-30 13:37
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
@Controller
public class FileUploadExcel {



    /**
     * @方法说明：单个或多个文件上传方法二
     * @时间： 2017-11-30 13:37
     * @创建人：wangl
     */
    @ResponseBody
    @RequestMapping(value = "/excel_load")
    public void excel(HttpServletResponse response, HttpServletRequest request) {
        File newFile = createNewFile();
        //File newFile = new File("C:/NGP-项目发布-紧急更新-V20180123.xlsx");

        // 新文件写入数据，并下载
        InputStream is = null;
        XSSFWorkbook workbook = null;
        XSSFSheet sheet = null;
        try {
            // 将excel文件转为输入流
            is = new FileInputStream(newFile);
            // 创建个workbook，
            workbook = new XSSFWorkbook(is);
            // 获取第一个sheet
            sheet = workbook.getSheetAt(0);
        } catch (Exception e1) {
            e1.printStackTrace();
        }

        if (sheet != null) {
            try {
                // 写数据
                FileOutputStream fos = new FileOutputStream(newFile);
                XSSFRow row = sheet.getRow(3);
                if (row == null) {
                    row = sheet.createRow(3);
                }
                XSSFCell cell = row.getCell(0);
                if (cell == null) {
                    cell = row.createCell(0);
                }

                // 定义一个list集合假数据
                List<Map<String, Object>> lst = new ArrayList();
                Map<String, Object> map1 = new HashMap<String, Object>();
                for (int i = 0; i < 10; i++) {
                    map1.put("id" + i, i);
                    lst.add(map1);
                }
                for (int m = 0; m < lst.size(); m++) {
                    Map<String, Object> map = lst.get(m);
                    row = sheet.createRow((int) m + 3);
                    for (int i = 0; i < 11; i++) {
                        String value = map.get("id" + m) + "";
                        if ("null".equals(value)) {
                            value = "0";
                        }
                        cell = row.createCell(i);
                        cell.setCellValue(value);
                    }

                }
                workbook.write(fos);
                fos.flush();
                fos.close();

                // 下载
                InputStream fis = new BufferedInputStream(new FileInputStream(newFile));
                //ServletActionContext.getResponse();

                byte[] buffer = new byte[fis.available()];
                fis.read(buffer);
                fis.close();
                response.reset();
                response.setContentType("text/html;charset=UTF-8");
                OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
                response.setContentType("application/x-msdownload");
                String newName = URLEncoder.encode("违法案件报表" + System.currentTimeMillis() + ".xlsx", "UTF-8");
                response.addHeader("Content-Disposition", "attachment;filename=\"" + newName + "\"");
                response.addHeader("Content-Length", "" + newFile.length());
                toClient.write(buffer);
                toClient.flush();
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                try {
                    if (null != is) {
                        is.close();
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public File createNewFile() {
        // 读取模板，并赋值到新文件************************************************************
        // 文件模板路径
        String path = ("C:/NGP-项目发布-紧急更新-V20180123.xlsx");
        File file = new File(path);
        // 保存文件的路径
        String realPath = ("D:/hy/项目发布");
        // 新的文件名
        String newFileName = "项目发布" + System.currentTimeMillis() + ".xlsx";
        // 判断路径是否存在
        File dir = new File(realPath);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        // 写入到新的excel
        File newFile = new File(realPath, newFileName);
        try {
            newFile.createNewFile();
            // 复制模板到新文件
            fileChannelCopy(file, newFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return newFile;
    }

    public void fileChannelCopy(File s, File t) {
        try {
            InputStream in = null;
            OutputStream out = null;
            try {
                in = new BufferedInputStream(new FileInputStream(s), 1024);
                out = new BufferedOutputStream(new FileOutputStream(t), 1024);
                byte[] buffer = new byte[1024];
                int len;
                while ((len = in.read(buffer)) != -1) {
                    out.write(buffer, 0, len);
                }
            } finally {
                if (null != in) {
                    in.close();
                }
                if (null != out) {
                    out.close();
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
