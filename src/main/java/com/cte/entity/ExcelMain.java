package com.cte.entity;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelMain {


    /**
     * 生成excel并下载
     */
    public void exportExcel() {

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
                for (int i = 0; i < 42; i++) {
                    map1.put("id" + i, i);
                    lst.add(map1);
                }
                for (int m = 0; m < lst.size(); m++) {
                    Map<String, Object> map = lst.get(m);
                    row = sheet.createRow((int) m + 3);
                    for (int i = 0; i < 42; i++) {
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
                HttpServletResponse response = null;
                byte[] buffer = new byte[fis.available()];
                fis.read(buffer);
                fis.close();
                response.reset();
                response.setContentType("text/html;charset=UTF-8");
                OutputStream toClient = new BufferedOutputStream(
                        response.getOutputStream());
                response.setContentType("application/x-msdownload");
                String newName = URLEncoder.encode(
                        "违法案件报表" + System.currentTimeMillis() + ".xlsx",
                        "UTF-8");
                response.addHeader("Content-Disposition",
                        "attachment;filename=\"" + newName + "\"");
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
        // 删除创建的新文件
        // this.deleteFile(newFile);
    }

    /**
     * 复制文件
     *
     * @param s 源文件
     * @param t 复制到的新文件
     */

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

    private String getSispPath() {
        /*String classPaths = IExportService.class.getResource("/").getPath();
        String[] aa = classPaths.split("/");
        String sispPath = "";
        for (int i = 1; i < aa.length - 2; i++) {
            sispPath += aa[i] + "/";
        }
        return sispPath;*/
        return "/log";
    }

    /**
     * 读取excel模板，并复制到新文件中供写入和下载
     *
     * @return
     */
    public File createNewFile() {
        // 读取模板，并赋值到新文件************************************************************
        // 文件模板路径
        String path = (getSispPath() + "uploadfile/违法案件报表.xlsx");
        File file = new File(path);
        // 保存文件的路径
        String realPath = (getSispPath() + "uploadfile");
        // 新的文件名
        String newFileName = "违法案件报表" + System.currentTimeMillis() + ".xlsx";
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

    /**
     * 下载成功后删除
     *
     * @param files
     */
    private void deleteFile(File... files) {
        for (File file : files) {
            if (file.exists()) {
                file.delete();
            }
        }
    }

}