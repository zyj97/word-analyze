package org.analyze.service;

import org.analyze.analyze.*;
import org.checkerframework.checker.units.qual.A;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

/**
 * @Author: zyj
 * @Date: 2021/5/31
 */
@Service
public class FileService {
    @Autowired
    DocExtract docExtract;
    @Autowired
    WordExtract wordExtract;
    @Autowired
    HWPFExtract hwpfExtract;
    @Autowired
    SpireExtract spireExtract;
    @Autowired
    TiKaExtract tiKaExtract;
    @Autowired
    DocxExtract docxExtract;

    public JSONObject analyze(MultipartFile file){
        JSONObject jsonObject = new JSONObject();
        String fileName = file.getOriginalFilename();
        if (fileName.endsWith("doc")){
            try {
//                docExtract.extract(file);
//                wordExtract.testReadByExtractor(file);
//                tiKaExtract.extract(file);
                jsonObject = hwpfExtract.testReadByDoc(file);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }else if(fileName.endsWith(".docx")){
            docxExtract.extract(file);


        }
        return  jsonObject;

    }


}
