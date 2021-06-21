package org.analyze.analyze;

import com.sun.scenario.effect.impl.sw.sse.SSEBlend_SRC_OUTPeer;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @Author: zyj
 * @Date: 2021/6/7
 */
@Component

public class DocxExtract {

    public void extract(MultipartFile file){
        try (InputStream inputStream = file.getInputStream()) {

            XWPFDocument doc = new XWPFDocument(inputStream);
            System.out.println("获取属性");
            POIXMLProperties poixmlProperties = doc.getProperties();
            POIXMLProperties.CustomProperties customProperties = poixmlProperties.getCustomProperties();


            POIXMLProperties.CoreProperties properties= doc.getProperties().getCoreProperties();
            System.out.println(properties.getCategory());
            System.out.println(properties.getContentStatus());
            System.out.println(properties.getContentType());


            List<XWPFParagraph> paras = doc.getParagraphs();
            for (XWPFParagraph paragraph : paras) {
                String a = paragraph.getText();
                if (StringUtils.isEmpty(a)){
                    continue;
                }
                a = a.replaceAll(" ","");
                if (StringUtils.isEmpty(a)){
                    continue;
                }

                List<XWPFRun> runs = paragraph.getRuns();
                String styleName = getDocxStyleName(doc,paragraph);
                String font = null;
                float fontSize = 0.0f;
                for (XWPFRun xwpfRun : runs) {
                    if (StringUtils.isEmpty(font)){
                        font = xwpfRun.getFontName();
                    }
                    if (fontSize == 0.0f){
                        fontSize = xwpfRun.getFontSize() +0.0f;
                    }
                }
                System.out.println("===============");
                System.out.println(a);
                System.out.println(styleName);
                System.out.println(paragraph.getStyle());
            }
    } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private String getDocxStyleName(XWPFDocument doc, XWPFParagraph paragraph){
        String styleName = "";
        int numStyles =doc.getStyles().getNumberOfStyles();

        String styleIndex = paragraph.getStyle();
        if (StringUtils.isEmpty(styleIndex)){
            return styleName;
        }


//        if (numStyles > styleIndex) {
            XWPFStyles style_sheet = doc.getStyles();

            XWPFStyle style = style_sheet.getStyle(styleIndex);

            styleName = style.getName();
            styleName.replaceAll(" ","");
            styleName.replaceAll("  ","");
            styleName.trim();
//        }
        return styleName;
    }
}
