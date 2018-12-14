package com.cte.controller;

import com.cte.entity.Message;
import com.cte.entity.Status;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;

/***
 *
 *
 * 描    述：文件上传
 *
 * 创 建 者： @author wl
 * 创建时间： 2018/12/13 18:07
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
public class FileUploadController {

    private static final Logger logger = LoggerFactory.getLogger(FileUploadController.class);

    /**
     * @方法说明：单个或多个文件上传方法二
     * @时间： 2017-11-30 13:37
     * @创建人：wangl
     */
    @ResponseBody
    @RequestMapping(value = "/uploadFile", method = RequestMethod.POST, produces = "application/json;charset=utf8")
    public Message uploadFileHandler(@RequestParam("file") MultipartFile file) throws IOException {
        //文件上传位置
        String path = "E:/upload/fileUpload/";

        logger.info("上传单个或多个文件");
        if (!file.isEmpty()) {
            InputStream is = null;

            FileOutputStream fos = null;
            int length = 0;
            try {
                long size = file.getSize();
                //使用Apache文件上传组件处理文件上传步骤：
                //1、创建一个DiskFileItemFactory工厂
                DiskFileItemFactory diskFileItemFactory = new DiskFileItemFactory();
                //2、创建一个文件上传解析器
                ServletFileUpload fileUpload = new ServletFileUpload(diskFileItemFactory);
                //解决上传文件名的中文乱码
                fileUpload.setHeaderEncoding("UTF-8");

                is = file.getInputStream();
                CreateDir.createDir(path);
                fos = new FileOutputStream(path + file.getOriginalFilename());
                byte[] buffer = new byte[(int) size];
                while ((length = is.read(buffer)) > 0) {
                    //使用FileOutputStream输出流将缓冲区的数据写入到指定的目录(savePath + "\\" + filename)当中
                    fos.write(buffer, 0, length);
                }
                Message msg = new Message();
                msg.setStatus(Status.SUCCESS);
                msg.setStatusMsg("File upload success");
                return msg;
            } catch (Exception e) {
                logger.error("上传单个或多个文件异常--->{}" + e);
                Message msg = new Message();
                msg.setStatus(Status.ERROR);
                msg.setError("File upload file");
                return msg;
            } finally {
                if (is != null) {
                    is.close();
                }
                if (fos != null) {
                    fos.close();
                }
            }
        } else {
            Message msg = new Message();
            msg.setStatus(Status.ERROR);
            msg.setError("File upload file");
            return msg;
        }
    }

    /**
     * @方法说明：单个或多个文件上传
     * @时间： 2017-11-30 13:37
     * @创建人：wangl
     */
    @ResponseBody
    @RequestMapping(value = "/uploadMultipleFile", method = RequestMethod.POST, produces = "application/json;charset=utf8")
    public Message uploadMultipleFileHandler(@RequestParam("file") MultipartFile[] files) throws IOException {

        Message msg = new Message();
        ArrayList<Integer> arr = new ArrayList<>();
        for (int i = 0; i < files.length; i++) {
            MultipartFile file = files[i];
            if (!file.isEmpty()) {
                InputStream in = null;
                OutputStream out = null;
                try {
                    String path = "/upload/fileUpload/" + file.getOriginalFilename();
                    CreateDir.createDir(path);

                    File serverFile = new File(path);
                    in = file.getInputStream();
                    out = new FileOutputStream(serverFile);

                    long size = file.getSize();
                    byte[] b = new byte[(int) size];
                    int len = 0;
                    while ((len = in.read(b)) > 0) {
                        out.write(b, 0, len);
                    }
                    logger.info("保存File路径--->{}" + serverFile.getAbsolutePath());
                } catch (Exception e) {
                    arr.add(i);
                } finally {
                    out.close();
                    in.close();
                }
            } else {
                arr.add(i);
            }
        }
        if (arr.size() > 0) {
            msg.setStatus(Status.ERROR);
            msg.setError("Files upload fail");
            msg.setErrorKys(arr);
        } else {
            msg.setStatus(Status.SUCCESS);
            msg.setStatusMsg("Files upload success");
        }
        return msg;
    }

}

