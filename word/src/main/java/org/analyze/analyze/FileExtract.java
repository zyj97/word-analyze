package org.analyze.analyze;

import org.springframework.web.multipart.MultipartFile;

/**
 * @Author: zyj
 * @Date: 2021/5/31
 */
public abstract class FileExtract {

    public abstract void extract(MultipartFile file);
}
