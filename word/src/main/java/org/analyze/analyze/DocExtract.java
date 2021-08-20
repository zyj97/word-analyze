package org.analyze.analyze;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author: zyj
 * @Date: 2021/5/31
 */
@Component
public class DocExtract {
    /**
     * 回车符ASCII码
     */
    private static final short ENTER_ASCII = 13;

    /**
     * 空格符ASCII码
     */
    private static final short SPACE_ASCII = 32;

    /**
     * 水平制表符ASCII码
     */
    private static final short TABULATION_ASCII = 9;

    public static String htmlText = "";
    public static String htmlTextTbl = "";
    public static int counter=0;
    public static int beginPosi=0;
    public static int endPosi=0;
    public static int beginArray[];
    public static int endArray[];
    public static String htmlTextArray[];
    public static boolean tblExist=false;

    //使用poi解析doc文档
    public void extract(MultipartFile file) throws IOException {
//        InputStream inputStream = null;
//        try{
//            inputStream = file.getInputStream();
//            POIFSFileSystem pfs = new POIFSFileSystem(inputStream);
//            HWPFDocument hwpf = new HWPFDocument(pfs);
//            Range range = hwpf.getRange();//得到文档的读取范围
//            TableIterator it = new TableIterator(range);
//            for (int o = 0; o< range.numCharacterRuns();o++){
//                String text = range.getCharacterRun(o).text();
//                text  = text.trim();
//                if (StringUtils.isEmpty(text)){
//                    continue;
//                }
//
//
//
//                short i = range.getCharacterRun(o).getStyleIndex();
//                System.out.println("============格式为" + i);
//                System.out.println(range.getCharacterRun(o).getFontName());
//                System.out.println(range.getCharacterRun(o).getFontSize());
//                System.out.println("颜色为" +range.getCharacterRun(o).getColor());
//                System.out.println("内容" +range.getCharacterRun(o).text());
//            }
//
//
//            if (it.hasNext()) {
//                Table table = it.next();
//                for (int row = 0; row < table.numRows(); row++) {
//                    TableRow rows = table.getRow(row);
//
//                    for (int column = 0; column < rows.numCells(); column++) {
//                        TableCell tableCell = rows.getCell(column);
//                        System.out.println("==================row");
//
//                        System.out.println("第" + (row + 1) + "行" + "第" + (column + 1) + "列" + "填写内容为： " + tableCell.text() + "======");
//
//                    }
//                }
//            }
//
//
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (Exception e) {
//            e.printStackTrace();
//        } finally {
//            inputStream.close();
//        }

        InputStream inputStream = file.getInputStream();
        POIFSFileSystem pfs = new POIFSFileSystem(inputStream);

        HWPFDocument doc = new HWPFDocument(pfs);
        // 文本
        Range range = doc.getRange();
        // 图片
        PicturesTable pTable = doc.getPicturesTable();

        int idx = 0;

        // 循环列表，处理段落内容
        for (int i = 0; i < range.numParagraphs(); i++) {
            // 段落内容取得
            Paragraph paragraph = range.getParagraph(i);
            if (pTable.hasPicture(paragraph.getCharacterRun(0))) {
                // 提取图片
                Picture pic = pTable.extractPicture(paragraph.getCharacterRun(0), false);
                // 返回POI建议的图片文件名
                String img = pic.suggestFullFileName();
                System.out.println(img);
                // 元素种类为段落的场合，读取段落内容
                Map<Integer, List<Object>> tmpMap = new HashMap<>();
                List<Object> tmpList = new ArrayList<>();
                tmpList.add(img);
                tmpMap.put(0, tmpList);
//                rsMap.put(idx + "-IMAGE", tmpMap);
            } else {
                if (!paragraph.isInTable()) {
                    // 元素种类为段落的场合，读取段落内容
                    Map<Integer, List<Object>> tmpMap = new HashMap<>();
                    readParagraph(paragraph);
//                    tmpMap.put(0, readParagraph(paragraph));
//                    rsMap.put(idx + "-PARAGRAPH", tmpMap);
                } else {
                    try {
                        // 元素种类为表格的场合，读取表格内容
                        Table table = range.getTable(paragraph);
                        readTable(table);
//                        rsMap.put(idx + "-TABLE", readTable(table));
                    } catch (IllegalArgumentException e) {
                        continue;
                    }
                }
            }
            idx++;
        }




    }

    void readTable(Table table){


    }

    void readParagraph(Paragraph paragraph){
        System.out.println(paragraph.text());

    }






}
