package com.study.exceluploader;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;


@RestController
public class ExcelController {

    @Autowired
    ExcelService excelService;

    @PostMapping("/excel")
    public void addExcel(@RequestBody MultipartFile file) throws Exception {
        excelService.readExcel(file);
    }
}
