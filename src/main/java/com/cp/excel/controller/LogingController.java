package com.cp.excel.controller;

import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * excel 操作controller
 *
 * @author cuipeng
 * @Create 2020-12-03-2:34 下午
 **/
@Slf4j
@Controller
public class LogingController {

    @RequestMapping(value="/")
    public String login(){
        return "excel1";
    }

}
