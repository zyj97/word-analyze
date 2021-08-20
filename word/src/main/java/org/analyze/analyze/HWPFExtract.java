package org.analyze.analyze;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.ListLevel;
import org.apache.poi.hwpf.model.ListTables;
import org.apache.poi.hwpf.model.StyleDescription;
import org.apache.poi.hwpf.model.StyleSheet;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
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


//        JSONArray text = printInfo(range,doc);
//        jsonObject.put("内容" ,text);

//        Range range1 =doc.getOverallRange();
//        printInfo(range1);
        printInfoAndTable(range,doc);








        //读表格
//        JSONObject table = readTable(range);
//        jsonObject.put("表格",table);
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
            String text = replaceSpecialChar( paragraph.text());
            if (StringUtils.isEmpty(text) || "\r".equals(text)){
               continue;
            }




            JSONObject jsonObject = new JSONObject();
            int runs = paragraph.numCharacterRuns();
            for (int r = 0;r<runs;r++){
                CharacterRun characterRun = paragraph.getCharacterRun(r);
                String fontName = characterRun.getFontName();
                int fontSize = characterRun.getFontSize();
            }
            String styleName = getStyleName(doc,paragraph);


            System.out.println("字体为================" +paragraph.getCharacterRun(0).getFontName());
            System.out.println("字号为================" +paragraph.getCharacterRun(0).getFontSize());
            System.out.println("格式为================" + styleName);
            jsonObject.put("格式", range.getParagraph(i).getStyleIndex());
            jsonObject.put("字体",range.getParagraph(i).getCharacterRun(0).getFontName());
            jsonObject.put("字号",range.getParagraph(i).getCharacterRun(0).getFontSize());
            jsonArray.add(jsonObject);
        }

       return jsonArray;
    }

    private String getStyleName(HWPFDocument doc,Paragraph paragraph){
        int numStyles =doc.getStyleSheet().numStyles();
        int styleIndex = paragraph.getStyleIndex();
        String styleName = "";
        if (numStyles > styleIndex) {
            StyleSheet style_sheet = doc.getStyleSheet();
            StyleDescription style = style_sheet.getStyleDescription(styleIndex);
            System.out.println( style.getBaseStyle());
            styleName = style.getName();
            styleName = styleName.replaceAll(" ","");
        }
        return  styleName;
    }

    private String replaceSpecialChar(String paragraphStr){
        int index3 = paragraphStr.indexOf(19);
        int index4 = paragraphStr.indexOf(20);

        while (index3 >=0 && index4 >=0){
            StringBuffer stringBuffer = new StringBuffer(paragraphStr);
            StringBuffer str = stringBuffer.replace(index3,index4+1 ,"");
            System.out.println("=====================");
            System.out.println(str.toString());
            paragraphStr = str.toString();
            index3 = paragraphStr.indexOf(19);
            index4 = paragraphStr.indexOf(20);
        }
        return paragraphStr;

    }


    private JSONArray printInfoAndTable(Range range,HWPFDocument doc) {
        //获取段落数
        int paraNum = range.numParagraphs();
        System.out.println(paraNum);
        JSONArray jsonArray = new JSONArray();
        boolean lastInTable = false;
        int tableNumber = 0;
        TableIterator tableIter = new TableIterator(range);

        for (int i=0; i<paraNum; i++) {

            Paragraph paragraph = range.getParagraph(i);


            if (paragraph.isInTable()){
                if (lastInTable){
                    continue;
                }
                readTableFromPara(range,tableNumber);
                tableNumber++;
                lastInTable = true;
                continue;
            }
            lastInTable = false;

//            System.out.println("======================================");
//
//            System.out.println(paragraph.getIndentFromRight());
//            System.out.println(paragraph.getFontAlignment());
//            System.out.println(paragraph.getLeftBorder());
            String text = replaceSpecialChar( paragraph.text());
            if (StringUtils.isEmpty(text) || "\r".equals(text)){
                continue;
            }
            int runs = paragraph.numCharacterRuns();
            for (int r = 0;r<runs;r++){
                CharacterRun characterRun = paragraph.getCharacterRun(r);
                String fontName = characterRun.getFontName();
                int fontSize = characterRun.getFontSize();

            }

            if (paragraph.isInList()) {
                ListTables lTable = doc.getListTables();
                ListLevel ll = lTable.getLevel(paragraph.getList().getLsid(), paragraph.getIlvl());

                int id = paragraph.getList().getLsid();
                int format = ll.getNumberFormat();
                String text1 = ll.getNumberText();
                System.out.println(String.valueOf(format));
                System.out.println(text1);
                System.out.println();

                StringBuffer sb = new StringBuffer(text);
                sb.insert(1,text1);
                System.out.println(sb.toString());
            }


//                    (StyleSheet styleSheet,
//                    ListTables listTables,
//            int ilfo)
            String styleName = getStyleName(doc,paragraph);
            System.out.println("内容为" + text);
        }

        return jsonArray;
    }



    private void readTableFromPara(Range range,int tableNumber){
        System.out.println("========================================================解析表格");
        //遍历range范围内的table。
        TableIterator tableIter = new TableIterator(range);
        Table table;
        TableRow row;
        TableCell cell;
        JSONObject  tableList  = new JSONObject();
        int i = -1;
        while (tableIter.hasNext()) {
            i++;
            table = tableIter.next();
            if (  i>tableNumber){
                break;
            }else if (i < tableNumber){
                continue;
            }

            JSONArray jsonArray = new JSONArray();


            int rowNum = table.numRows();
            for (int j=0; j<rowNum; j++) {
                row = table.getRow(j);
                int cellNum = row.numCells();
                for (int k=0; k<cellNum; k++) {
                    cell = row.getCell(k);
                    //输出单元格的文本
                    JSONObject tableJson = new JSONObject();
                    System.out.println(i+ "=====行号是"+j + "=====列号" + k +"表格内容是" + cell.text());
                    tableJson.put("行号",j);
                    tableJson.put("列号",k);
                    tableJson.put("内容",cell.text());
                    jsonArray.add(tableJson);

                }
            }
            tableList.put("表格" + i,jsonArray);

        }






    }


//    FileInputStream fis = new FileInputStream(file);
//    HWPFDocument doc = new HWPFDocument(fis);
//    // 文本
//    Range range = doc.getRange();
//    // 图片
//    PicturesTable pTable = doc.getPicturesTable();
//
//    int idx = 0;
//
//    // 循环列表，处理段落内容
//        for (int i = 0; i < range.numParagraphs(); i++) {
//        // 段落内容取得
//        Paragraph paragraph = range.getParagraph(i);
//        if (pTable.hasPicture(paragraph.getCharacterRun(0))) {
//            // 提取图片
//            Picture pic = pTable.extractPicture(paragraph.getCharacterRun(0), false);
//            // 返回POI建议的图片文件名
//            String img = pic.suggestFullFileName();
//            // 元素种类为段落的场合，读取段落内容
//            Map<Integer, List<Object>> tmpMap = new HashMap<>();
//            List<Object> tmpList = new ArrayList<>();
//            tmpList.add(img);
//            tmpMap.put(0, tmpList);
//            rsMap.put(idx + "-IMAGE", tmpMap);
//        } else {
//            if (!paragraph.isInTable()) {
//                // 元素种类为段落的场合，读取段落内容
//                Map<Integer, List<Object>> tmpMap = new HashMap<>();
//                tmpMap.put(0, readParagraph(paragraph));
//                rsMap.put(idx + "-PARAGRAPH", tmpMap);
//            } else {
//                try {
//                    // 元素种类为表格的场合，读取表格内容
//                    Table table = range.getTable(paragraph);
//                    rsMap.put(idx + "-TABLE", readTable(table));
//                } catch (IllegalArgumentException e) {
//                    continue;
//                }
//            }
//        }
//        idx++;
//    }


}
