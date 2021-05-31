package org.analyze.service;

import org.analyze.analyze.DocExtract;
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

    public void analyze(MultipartFile file){
        String fileName = file.getOriginalFilename();
        if (fileName.endsWith("doc")){
            try {
                docExtract.extract(file);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }


}
