package com.cp.excel.controller;

import com.cp.excel.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * excel 操作controller
 *
 * @author cuipeng
 * @Create 2020-12-03-2:34 下午
 **/
@Slf4j
@Controller
@RequestMapping("/excel")
public class ExcelController {

    @RequestMapping(value="addAll",method= RequestMethod.POST)
    @ResponseBody
    public String addAll(String compareCols , String addCols, String templateName ,MultipartFile template, List<MultipartFile> addFiles) throws Exception {
        List<Integer> addColList = convertStrToList(addCols);
        String TemplatePathName = ExcelUtil.getCopyTemplate(template,templateName+System.currentTimeMillis(), addColList);
        List<Integer> compareColList = convertStrToList(compareCols);
        List<Integer> allcolList = ExcelUtil.getAllColsAddOne(addColList);
        File tempalterFile = new File(TemplatePathName);
        //按照名称排序
        addFiles =  addFiles.stream().sorted((p1,p2)->{
            return p1.getOriginalFilename().compareTo(p2.getOriginalFilename());
        }).collect(Collectors.toList());

        for(int i=0;i<addFiles.size();i++){
            ExcelUtil.addTemplateData(tempalterFile,addFiles.get(i),compareColList,addColList,allcolList);
            allcolList = ExcelUtil.getAllColsAddOne(allcolList);
        }
        ExcelUtil.clearUpTemplate(tempalterFile);
        return TemplatePathName;
    }



    private List<Integer> convertStrToList(String str){
        String[] strArray = str.split(",");
        List<Integer> list = new ArrayList<>();
        for (String s : strArray) {
            list.add(Integer.parseInt(s));
        }
        return list;
    }
}
