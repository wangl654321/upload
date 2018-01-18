package com.cte.controller;

import com.cte.entity.ImportExcelUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/***
 *
 *
 * 描    述：文件上传
 *
 * 创 建 者： wangl
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

    private static final Logger logger = LoggerFactory.getLogger(FileUploadController.class);

    @Autowired
    private ImportExcelUtil importExcelUtil;
    /**
     * @方法说明：单个或多个文件上传方法二
     * @时间： 2017-11-30 13:37
     * @创建人：wangl
     */
    @ResponseBody
    @RequestMapping(value = "/uploadFile_file",method = RequestMethod.POST)
    public String uploadFileHandler(@RequestParam("file") MultipartFile file) throws Exception {

        Map<String, String> map = new HashMap<>();
        map.put("id", "id");
        map.put("姓名", "name");
        map.put("年龄", "age");
        map.put("时间", "time");
        map.put("金额", "money");
        List<Map<String, Object>> ls = importExcelUtil.parseExcel(file, map);
        logger.info(ls.toString());

        return "ok";
    }

}

