package org.analyze.analyze;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @Author: zyj
 * @Date: 2021/6/3
 */
@Component
public class HWPFExtract {
    /** 分隔符数组 */
    private final static String[] separaters = {"。", "，", "！", "？", "；"};

    public JSONObject testReadByDoc(MultipartFile file) throws Exception {
        JSONObject jsonObject = new JSONObject();
        InputStream is = file.getInputStream();
        HWPFDocument doc = new HWPFDocument(is);
        //输出书签信息
//        this.printInfo(doc.getBookmarks());
        //输出文本
//        System.out.println(doc.getDocumentText());
        Range range = doc.getRange();
//        Range range = doc.getCommentsRange();
////    this.insertInfo(range);


        JSONArray text = printInfo(range,doc);
        jsonObject.put("内容" ,text);

//        Range range1 =doc.getOverallRange();
//        printInfo(range1);








        //读表格
        JSONObject table = readTable(range);
        jsonObject.put("表格",table);
        //读列表
//        this.readList(range);
        //删除range
//        Range r = new Range(2, 5, doc);
//        r.delete();//在内存中进行删除，如果需要保存到文件中需要再把它写回文件
        //把当前HWPFDocument写到输出流中
//        doc.write(new FileOutputStream("D:\\test.doc"));
//        readParagraph(range);
        this.closeStream(is);
        return jsonObject;
    }




    /**
     * 关闭输入流
     * @param is
     */
    private void closeStream(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 输出书签信息
     * @param bookmarks
     */
    private void printInfo(Bookmarks bookmarks) {
        int count = bookmarks.getBookmarksCount();
        System.out.println("书签数量：" + count);
        Bookmark bookmark;
        for (int i=0; i<count; i++) {
            bookmark = bookmarks.getBookmark(i);
            System.out.println("书签" + (i+1) + "的名称是：" + bookmark.getName());
            System.out.println("开始位置：" + bookmark.getStart());
            System.out.println("结束位置：" + bookmark.getEnd());
        }
    }

    /**
     * 读表格
     * 每一个回车符代表一个段落，所以对于表格而言，每一个单元格至少包含一个段落，每行结束都是一个段落。
     * @param range
     */
    private JSONObject readTable(Range range) {
        System.out.println("========================================================解析表格");
        //遍历range范围内的table。
        TableIterator tableIter = new TableIterator(range);
        Table table;
        TableRow row;
        TableCell cell;
        JSONObject  tableList  = new JSONObject();
        int i = 0;
        while (tableIter.hasNext()) {
            i++;
            JSONArray jsonArray = new JSONArray();
            if (i>1){
                break;
            }


            table = tableIter.next();
            int rowNum = table.numRows();
            for (int j=0; j<rowNum; j++) {
                row = table.getRow(j);
                int cellNum = row.numCells();
                for (int k=0; k<cellNum; k++) {
                    cell = row.getCell(k);
                    //输出单元格的文本
                    JSONObject tableJson = new JSONObject();
                    System.out.println("=====行号是"+j + "=====列号" + k +"表格内容是" + cell.text());
                    tableJson.put("行号",j);
                    tableJson.put("列号",k);
                    tableJson.put("内容",cell.text());
                    jsonArray.add(tableJson);

                }
            }
            tableList.put("表格" + i,jsonArray);

        }
        return tableList;

    }

    /**
     * 读列表
     * @param range
     */
    private void readList(Range range) {
        int num = range.numParagraphs();
        Paragraph para;
        for (int i=0; i<num; i++) {
            para = range.getParagraph(i);
            if (para.isInList()) {
                System.out.println("list: " + para.text());
            }
        }
    }




    /**
     * 输出Range
     * @param range
     */
    private JSONArray printInfo(Range range,HWPFDocument doc) {
        //获取段落数
        int paraNum = range.numParagraphs();
        System.out.println(paraNum);
        JSONArray jsonArray = new JSONArray();
        for (int i=0; i<paraNum; i++) {

            Paragraph paragraph = range.getParagraph(i);
            //文本在表格中 所以不需要高
            if (paragraph.isInTable()){
                continue;
            }
            String text = paragraph.text().trim();

            if (StringUtils.isEmpty(text)){
               continue;
            }
            if (text== "\r"){
                continue;
            }



            JSONObject jsonObject = new JSONObject();
            String paragraphStr = "段落" + (i+1) + "：" + text;
            int runs = paragraph.numCharacterRuns();
            for (int r = 0;r<runs;r++){
                CharacterRun characterRun = paragraph.getCharacterRun(r);
                String fontName = characterRun.getFontName();
                int fontSize = characterRun.getFontSize();
//                if (StringUtils.isEmpty(fontName) && fontSize != 0){
//                    break;
//                }


            }
            int numStyles =doc.getStyleSheet().numStyles();
            int styleIndex = paragraph.getStyleIndex();
            String styleName = "";
            if (numStyles > styleIndex) {
                StyleSheet style_sheet = doc.getStyleSheet();
                StyleDescription style = style_sheet.getStyleDescription(styleIndex);
                styleName = style.getName();
                styleName = styleName.replaceAll(" ","");
            }

//            System.out.println("");


            System.out.println("字体为================" +paragraph.getCharacterRun(0).getFontName());
            System.out.println("字号为================" +paragraph.getCharacterRun(0).getFontSize());
            System.out.println("格式为================" + styleName);
            System.out.println("内容为" + paragraphStr);
//            System.out.println("字体为================" + range.getCharacterRun(i).getStyleIndex());
//            System.out.println("格式为================" + range.getParagraph(i).getStyleIndex());

            if ("标题 1".equals(styleName)){
                System.out.println("000000000000000000000");
            }
            if ("标题1".equals(styleName)){
                System.out.println("444444444444444444444444444444444444");
            }
            jsonObject.put("格式", range.getParagraph(i).getStyleIndex());
            jsonObject.put("字体",range.getParagraph(i).getCharacterRun(0).getFontName());
            jsonObject.put("字号",range.getParagraph(i).getCharacterRun(0).getFontSize());
            jsonObject.put("内容","段落" + (i+1) + "：" +text );
            jsonArray.add(jsonObject);
        }

       return jsonArray;
    }

    /**
     * 插入内容到Range，这里只会写到内存中
     * @param range
     */
    private void insertInfo(Range range) {
        range.insertAfter("Hello");
    }



}
