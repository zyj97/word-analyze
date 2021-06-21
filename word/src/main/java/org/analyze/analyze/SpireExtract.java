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
            Document document = new Document();


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
