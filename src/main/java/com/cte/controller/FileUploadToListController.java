package com.cte.controller;

import com.cte.entity.ExcelUtils;
import com.cte.entity.ImportExcelUtil;
import org.apache.poi.hssf.usermodel.*;
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
import java.text.SimpleDateFormat;
import java.util.*;

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
public class FileUploadToListController {

    private static final Logger logger = LoggerFactory.getLogger(FileUploadToListController.class);

    /**
     * @方法说明：单个或多个文件上传方法二
     * @时间： 2017-11-30 13:37
     * @创建人：wangl
     */
    @ResponseBody
    @RequestMapping(value = "/uploadFile_file", method = RequestMethod.POST)
    public String uploadFileHandler(@RequestParam("file") MultipartFile file) throws Exception {
        //请看document中的load.xlxs的模板
        Map<String, String> map = new HashMap<>();
        map.put("id", "id");
        map.put("姓名", "name");
        map.put("年龄", "age");
        map.put("时间", "time");
        map.put("金额", "money");
        List<Map<String, Object>> ls = ImportExcelUtil.parseExcel(file, map);
        System.out.println(ls.toString());

        return "ok";
    }

    /**
     * @方法说明：单个或多个文件上传方法二
     * @时间： 2017-11-30 13:37
     * @创建人：wangl
     */
    @ResponseBody
    @RequestMapping(value = "/excel")
    public void excel(HttpServletResponse response, HttpServletRequest request) {


        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String format = simpleDateFormat.format(new Date()).replace("00:00:00.0", "").replace(".0", "");

        String[] queryName = new String[5];
        queryName[0] = "紧急更新时间 " + format;
        queryName[1] = "紧急更新时间 " + "紧急更新";
        queryName[2] = "当值更新负责人 " + "侯建春";
        queryName[3] = " ";
        queryName[4] = "填写时间 " + format;

        List<Object[]> contents = new ArrayList<>();

        String[] content = new String[15];

        content[0] = "1";
        content[1] = "manage";
        content[2] = "manage_web";
        content[3] = "hotfix/201801/20180123_roule";
        content[4] = "对账规则二代添加支持添加特殊的结束标记";
        content[5] = "";
        content[6] = "王路";
        content[7] = "王路";
        content[8] = "李金徽";
        content[9] = "王路";
        content[10] = "李金徽";
        content[11] = "";
        content[12] = "";
        content[13] = "";
        content[13] = "";
        content[13] = "";
        contents.add(content);

        String[] headers = new String[]{"序号", "jenkins/job", "部署服务web-server", "hotfix分支", "变更内容说明", "数据库",
                "开发负责人", "开发负责人", "项目负责人", "项目负责人", "需求方", "预发布检查结果", "正式检查结果", "检查人", "备注"};

        try {
            ExcelUtils excelUtils = new ExcelUtils("D:/hy/NGP-项目发布-紧急更新.xls");

            excelUtils.createSheet("NGP-项目发布-紧急更新", "生产环境服务紧急更新申请单", queryName, headers, contents);
            excelUtils.excelFinal();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
