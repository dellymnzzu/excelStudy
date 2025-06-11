package com.study.controller;

import com.study.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Controller
public class ExcelController {

    @Autowired
    private ExcelService excelService;
    @PostMapping("/upload")
    public String upload(@RequestParam("file") MultipartFile file) throws IOException {
        excelService.readExcel(file);
        return "redirect:/";
    }
    @GetMapping("/")
    public String index() {
        return "upload";
    }
}
