package org.analyze.service;

import org.analyze.analyze.*;
import org.analyze.utils.MultipartFileToFileUtils;
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
    @Value(value = "${path.savePath}")
    private String savePath;

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
//            docxExtract.extract(file);
            try {
                xwpfWordExtract.testReadByExtractor(file);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return  jsonObject;

    }

    public String zipDecompression(MultipartFile file) throws Exception {

        String inputFile = file.getOriginalFilename();

        File srcFile = MultipartFileToFileUtils.multipartFileToFile(file);
        // 判断源文件是否存在
        if (!srcFile.exists()) {
            throw new Exception(srcFile.getPath() + "所指文件不存在");
        }
        String destDirPath = savePath.concat(File.separator).concat(inputFile.replace(".zip", ""));
        //创建压缩文件对象
        ZipFile zipFile = new ZipFile(srcFile);
        //开始解压
        Enumeration<?> entries = zipFile.entries();
        while (entries.hasMoreElements()) {
            ZipEntry entry = (ZipEntry) entries.nextElement();
            // 如果是文件夹，就创建个文件夹
            if (entry.isDirectory()) {
                srcFile.mkdirs();
            } else {
                // 如果是文件，就先创建一个文件，然后用io流把内容copy过去
                File targetFile = new File(savePath + "/" + entry.getName());
                // 保证这个文件的父文件夹必须要存在
                if (!targetFile.getParentFile().exists()) {
                    targetFile.getParentFile().mkdirs();
                }
                targetFile.createNewFile();
                // 将压缩文件内容写入到这个文件中
                InputStream is = zipFile.getInputStream(entry);
                FileOutputStream fos = new FileOutputStream(targetFile);
                int len;
                byte[] buf = new byte[1024];
                while ((len = is.read(buf)) != -1) {
                    fos.write(buf, 0, len);
                }
                // 关流顺序，先打开的后关闭
                fos.close();
                is.close();
            }
        }
        return  "11";
    }


}
