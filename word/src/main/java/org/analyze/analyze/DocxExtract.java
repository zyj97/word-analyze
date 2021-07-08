package org.analyze.analyze;

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

    public void extract(MultipartFile file) {
        try (InputStream inputStream = file.getInputStream()) {

            XWPFDocument doc = new XWPFDocument(inputStream);
            System.out.println("获取属性");

            POIXMLProperties.CoreProperties properties = doc.getProperties().getCoreProperties();
            List<IBodyElement> lstElement = doc.getBodyElements();
            for (int i = 0; i < lstElement.size(); i++) {
                if (BodyElementType.PARAGRAPH.equals(lstElement.get(i).getElementType())) {
                    XWPFParagraph paragraph = (XWPFParagraph)lstElement.get(i);
                    analyzeParagraph(paragraph,doc);
                } else if (BodyElementType.TABLE.equals(lstElement.get(i).getElementType())) {
                    XWPFTable table = (XWPFTable)lstElement.get(i);
                    analyzeTable(table);
                }else if (BodyElementType.CONTENTCONTROL.equals(lstElement.get(i).getElementType())){
                    System.out.println("不知道是什么");
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void analyzeParagraph(XWPFParagraph paragraph,XWPFDocument doc){
        String a = paragraph.getText();
        if (StringUtils.isEmpty(a)) {
            return;
        }
        a = a.replaceAll(" ", "");
        if (StringUtils.isEmpty(a)) {
            return;
        }

        List<XWPFRun> runs = paragraph.getRuns();
        String styleName = getDocxStyleName(doc, paragraph);
        String font = null;
        float fontSize = 0.0f;
        for (XWPFRun xwpfRun : runs) {
            if (StringUtils.isEmpty(font)) {
                font = xwpfRun.getFontName();
            }
            if (fontSize == 0.0f) {
                fontSize = xwpfRun.getFontSize() + 0.0f;
            }
        }
        System.out.println("===============");
        System.out.println(a);
        System.out.println(styleName);
    }

    private String getDocxStyleName(XWPFDocument doc, XWPFParagraph paragraph) {
        String styleName = "";
        int numStyles = doc.getStyles().getNumberOfStyles();

        String styleIndex = paragraph.getStyle();
        if (StringUtils.isEmpty(styleIndex)) {
            return styleName;
        }


//        if (numStyles > styleIndex) {
        XWPFStyles style_sheet = doc.getStyles();

        XWPFStyle style = style_sheet.getStyle(styleIndex);

        styleName = style.getName();
        styleName.replaceAll(" ", "");
        styleName.replaceAll("  ", "");
        styleName.trim();
//        }
        return styleName;
    }

    private void analyzeTable(XWPFTable table){


        List<XWPFTableRow> rows = table.getRows();
        for (int row = 0; row < rows.size(); row++) {
            XWPFTableRow xwpfTableRow = rows.get(row);
            // 获取表格的每个单元格
            List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
            for (int column = 0; column < tableCells.size(); column++) {
                XWPFTableCell cell = tableCells.get(column);
                // 获取单元格的内容
                String text1 = cell.getText();
                System.out.println("=========是表格====" + row+"===============" + column);
                System.out.println(text1);

            }
        }

    }
}
