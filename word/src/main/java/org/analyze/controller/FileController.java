package org.analyze.controller;

import org.analyze.analyze.DocExtract;
import org.analyze.pojo.ResMessage;
import org.analyze.service.FileService;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

/**
 * @Author: zyj
 * @Date: 2021/5/31
 */
@RestController
@RequestMapping("/file")
public class FileController {
    @Autowired
    private FileService fileService;


    @PostMapping("/upload")
    public JSONObject upload(MultipartFile file){
        return fileService.analyze(file);

    }


}
