package org.analyze.service;

import org.analyze.analyze.*;
import org.analyze.utils.MultipartFileToFileUtils;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * @Author: zyj
 * @Date: 2021/5/31
 */
@Service
public class FileService {
    @Autowired
    DocExtract docExtract;
    @Autowired
    HWPEWordExtract HWPEWordExtract;
    @Autowired
    HWPFExtract hwpfExtract;
    @Autowired
    TiKaExtract tiKaExtract;
    @Autowired
    DocxExtract docxExtract;
    @Autowired
    XWPFWordExtract xwpfWordExtract;
    @Autowired
    PDFExtract pdfExtract;
    @Autowired
    DocxExtract extract;

    @Value(value = "${path.savePath}")
    private String savePath;

    public JSONObject analyze(MultipartFile file){
        JSONObject jsonObject = new JSONObject();
        String fileName = file.getOriginalFilename();
        JSONArray tables= new JSONArray();

        if (fileName.endsWith("doc")){
            try {
                docExtract.extract(file);
//                wordExtract.testReadByExtractor(file);
//                tiKaExtract.extract(file);
//                 hwpfExtract.testReadByDoc(file);
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }else if(fileName.endsWith(".docx")){
//            docxExtract.extract(file);
            try {
                extract.extract(file);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }else if (fileName.endsWith(".pdf")){

            pdfExtract.extract(file);
        }
        return  jsonObject;

    }

    public String zipDecompression(MultipartFile file) throws Exception {

        String inputFile = file.getOriginalFilename();

        File srcFile = MultipartFileToFileUtils.multipartFileToFile(file);
        // ???????????????????????????
        if (!srcFile.exists()) {
            throw new Exception(srcFile.getPath() + "?????????????????????");
        }
        String destDirPath = savePath.concat("1").concat(File.separator).concat(inputFile.replace(".zip", ""));
        //????????????????????????
        ZipFile zipFile = new ZipFile(srcFile);
        //????????????
        Enumeration<?> entries = zipFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = (ZipEntry) entries.nextElement();
            // ??????????????????????????????????????????
            if (entry.isDirectory()) {
                srcFile.mkdirs();
            } else {
                // ??????????????????????????????????????????????????????io????????????copy??????
                File targetFile = new File( savePath.concat(File.separator).concat("1").concat(File.separator).concat(entry.getName()) );
                // ????????????????????????????????????????????????
                if (!targetFile.getParentFile().exists()) {
                    targetFile.getParentFile().mkdirs();
                }
                targetFile.createNewFile();
                // ?????????????????????????????????????????????
                InputStream is = zipFile.getInputStream(entry);
                FileOutputStream fos = new FileOutputStream(targetFile);
                int len;
                byte[] buf = new byte[1024];
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
                // ????????????????????????????????????
                fos.close();
                is.close();
            }
        }
        return  "11";
    }


}
