package org.analyze.analyze;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.collections.ParagraphCollection;
import com.spire.doc.collections.SectionCollection;
import com.spire.doc.documents.Paragraph;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author: zyj
 * @Date: 2021/6/4
 *
 *
 * spire有限制 试试删了可以不
 */
@Component
public class SpireExtract {
    public void testReadByExtractor() throws Exception {

        try  {
            int paragraphsCount = 0;
//            Document document = new Document("");
//            SectionCollection sections = document.getSections();
//            document.getSections().get(0).getParagraphs().removeAt(1);
//            document.saveToFile("D:\\"+ paragraphsCount +"RemoveParagraph.doc",FileFormat.Doc);

            Document document = new Document();

            //加载示例文档
            document.loadFromFile("D:\\2018B2(2)(1)(1)(1)(3)(1).doc");

            //删除第一节的第二段
            document.getSections().get(0).getParagraphs().removeAt(1);

            //保存文档
            document.saveToFile("D:\\"+ paragraphsCount +"RemoveParagraph.doc", FileFormat.Docx_2013);
//            //一个section解析一次
//            System.out.println(sections.getCount());
//            for (int i = 0; i < sections.getCount(); i++) {
//
//                Section section = sections.get(i);
//               paragraphsCount = paragraphsCount + section.getParagraphs().getCount();
//                System.out.println(paragraphsCount);
//            }
//
//            System.out.println(paragraphsCount);
//
//            for (int i = 0; i < sections.getCount(); i++) {
//                Section section = sections.get(i);
//                ParagraphCollection paragraphs = section.getParagraphs();
//
//                int max = 480;
//                if ( max > paragraphs.getCount()){
//                    max = paragraphs.getCount();
//                }
//
////                for (int j = 0;j<max;j++) {
////                    Paragraph paragraph = paragraphs.get(j);
////                    String styleName = paragraph.getStyleName();
////                    String paragraphText = paragraph.getText();
////                    System.out.println("++++++"  + j  );
////                    System.out.println(paragraphText);
//////
////                }
//
////                for (int j = 0;j<200;j++){
//
////                }
//
//            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
