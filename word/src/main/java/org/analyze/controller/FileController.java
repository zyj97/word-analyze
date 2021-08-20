package org.analyze.controller;

import org.analyze.analyze.DocExtract;
import org.analyze.analyze.XWPFWordExtract;
import org.analyze.pojo.ResMessage;
import org.analyze.service.FileService;
import org.analyze.service.GenerateMdFile;
import org.apache.commons.lang3.StringUtils;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
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
    @Autowired
    XWPFWordExtract xwpfWordExtract;
    @Autowired
    GenerateMdFile generateMdFile;


    @PostMapping("/upload")
    public JSONObject upload(MultipartFile file){
        return fileService.analyze(file);

    }


    @PostMapping("/zipUncompress")
    public String zipUncompress(MultipartFile file){
        try {
            return fileService.zipDecompression(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "22";

    }

//    @PostMapping("/md")
//    public String generateMdFile(MultipartFile file){
//
//    }


    @PostMapping("/mdStr")
    public String generateMdFile(MultipartFile file1) throws Exception {
//        String str = xwpfWordExtract.testReadByExtractor(file1);
        String result = generateMdFile.ganderMdfile(file1);

        return result;


    }




}
