package com.study.exceluploader;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Map;

@Service
public class ExcelService {

    @Autowired
    ExcelUtil excelUtil;

    public void readExcel(MultipartFile file) throws Exception {

        List<Map<String, Object>> listMap = excelUtil.getListData(file);
    }
}
