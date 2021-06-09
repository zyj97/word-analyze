package org.analyze.analyze;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
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
            Integer tableCount = 0;

//            List<XWPFTable> tables = doc.getTables();
//            for (XWPFTable table : tables) {
//                // 获取表格的行
//                if (tableCount >0){
//                    break;
//                }
//
//                List<XWPFTableRow> rows = table.getRows();
//                for (int row = 0;row<rows.size() ;row ++ ) {
//                    XWPFTableRow xwpfTableRow = rows.get(row);
//                    // 获取表格的每个单元格
//                    List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
//                    for (int column=0;column<tableCells.size();column++ ) {
//                        XWPFTableCell cell  = tableCells.get(column);
//                        // 获取单元格的内容
//                        String text1 = cell.getText();
//                        System.out.println(text1);
//                        HashMap<Integer,String> tableInfo = tableInfoMap.get(row + 1);
//                        if (tableInfo == null){
//                            tableInfo = new HashMap<>();
//                        }
//                        tableInfo.put(column + 1,text1);
//                        tableInfoMap.put(row +1,tableInfo);
//                    }
//                }
//                tableCount ++;
//            }

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
                String styleName = paragraph.getStyle();
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
            }
    } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
