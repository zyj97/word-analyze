package org.analyze.utils;


import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;

public class MultipartFileToFileUtils {
    private static final Logger logger = LoggerFactory.getLogger(MultipartFileToFileUtils.class);
    public static File multipartFileToFile(MultipartFile file) throws Exception {
        File toFile = null;
        if (file.equals("") || file.getSize() <= 0) {
            file = null;
        } else {
            InputStream ins = null;
            ins = file.getInputStream();
            toFile = new File(file.getOriginalFilename());
            inputStreamToFile(ins, toFile);
            ins.close();
        }
        return toFile;
    }


    //获取流文件
    public static void inputStreamToFile(InputStream ins, File file) throws IOException {
        OutputStream os = new FileOutputStream(file);
        try {

            int bytesRead = 0;
            byte[] buffer = new byte[8192];
            while ((bytesRead = ins.read(buffer, 0, 8192)) != -1) {
                os.write(buffer, 0, bytesRead);
            }

        } catch (Exception e) {
            logger.info("写入失败" + e.getMessage());
        }finally {
            os.close();
            ins.close();
        }
    }

    /**
     * 删除本地临时文件
     * @param file
     */
//    public static void delteTempFile(File file) {
//        if (file != null) {
//            File del = new File(file.toURI());
//            del.delete();
//        }
//    }

}
