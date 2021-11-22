package org.analyze.analyze;

import org.analyze.pojo.TableContent;
import org.analyze.utils.WordGradeExtractUtils;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.json.simple.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.*;

/**
 * @Author: zyj
 * @Date: 2021/5/31
 */
@Component
public class DocExtract {
    @Autowired
    WordGradeExtractUtils wordGradeExtractUtils;
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

        HashMap<Integer, List<Integer>> grade = new HashMap<>();

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
                    readParagraph(doc,paragraph,grade);
//                    tmpMap.put(0, readParagraph(paragraph));
//                    rsMap.put(idx + "-PARAGRAPH", tmpMap);
                } else {
                    try {
                        // 元素种类为表格的场合，读取表格内容
                        Table table = range.getTable(paragraph);
                        Object o = readTable(table);
//                        rsMap.put(idx + "-TABLE", readTable(table));
                    } catch (IllegalArgumentException e) {
                        continue;
                    }
                }
            }
            idx++;
        }




    }

    /**
     * 解析doc格式表格
     * @param table
     */
    Object readTable(Table table){

//        TableIterator tableIter = new TableIterator(range);
        HashMap<Integer, HashMap<Integer, HashMap<Integer, TableContent>>> tables = new HashMap<>();

        TableRow tableRow;

        //最到单元格数量
        int maxCellNum = 0;

        Map<Integer,Integer> widthList = new HashMap();

            int rowNum = table.numRows();
            HashMap<Integer, HashMap<Integer, TableContent>> tableInfo = new HashMap();

            for (int j = 0; j < rowNum; ++j) {
                TableRow row = table.getRow(j);
                int cellNum = row.numCells();
                HashMap<Integer, TableContent> rowInfo = new HashMap();

                //如果单元格数量大于最打单元格数量 就认定该行为子单位
                boolean updateTableData = cellNum > maxCellNum;
                if ( updateTableData){
                    maxCellNum = cellNum;
                }
                for (int k = 0; k < cellNum; ++k) {
                    TableCell cell = row.getCell(k);
                    if (updateTableData){
                        widthList.put(k,cell.getWidth());
                    }


                    if (cell.isVerticallyMerged() && !cell.isFirstVerticallyMerged()){
                        String text = "VMERGE";
                        TableContent tableContent = new TableContent();
                        tableContent.setTableWidth(cell.getWidth());
                        tableContent.setRowId(j + 1);
                        tableContent.setContent(text);
                        tableContent.setTitle("");
                        tableContent.setColumnId(k + 1);
                        rowInfo.put(k + 1, tableContent);

                    }else if(cell.isMerged()){
                        //横向合并
                        String text = "HMERGE";
                        TableContent tableContent = new TableContent();
                        tableContent.setRowId(j + 1);
                        tableContent.setContent(text);
                        tableContent.setTitle("");
                        tableContent.setColumnId(k + 1);
                        tableContent.setTableWidth(cell.getWidth());
                        rowInfo.put(k + 1, tableContent);

                    }else {
                        String text = cell.text();
                        text = text.replaceAll("/t", "");
                        TableContent tableContent = new TableContent();
                        tableContent.setRowId(j + 1);
                        tableContent.setContent(text);
                        tableContent.setTitle("");
                        tableContent.setColumnId(k + 1);
                        tableContent.setTableWidth(cell.getWidth());
                        rowInfo.put(k + 1, tableContent);
                    }


                }
                tableInfo.put(j + 1, rowInfo);
            }


        HashMap<Integer, HashMap<Integer, TableContent>> tableInfoNew = new HashMap();
        Set<Integer> rowList = tableInfo.keySet();
        for (Integer row : rowList){
            HashMap<Integer, TableContent> cellMap = tableInfo.get(row);
            HashMap<Integer, TableContent> cellNewMap = new HashMap<>();
            Set<Integer> cellKey = cellMap.keySet();

            //如果单元格数量等于 一行最大单元格数量 可认为该行没有进行合并 迭代下个单元格
            if (cellKey.size() >=  maxCellNum){
                tableInfoNew.put(row,cellMap);
                continue;
            }
            int cellWidthTotal = 0;
            int maxCellWidthTotal = 0;
            int maxCellIndex = 0;


            //当前单元格的列号
            int currentColumnNumber = 1;



            for (Integer cellId:cellKey){

                //已经合并的表格数量 只是合并的单元格
                int merged = 0;
                TableContent tableContent = cellMap.get(cellId);
                String title = tableContent.getTitle();
                int width = tableContent.getTableWidth();
                cellWidthTotal = cellWidthTotal + width;
                if (maxCellIndex < widthList.size()){
                    maxCellWidthTotal = maxCellWidthTotal + widthList.get(maxCellIndex);
                    maxCellIndex ++;
                }


                //
                while (cellWidthTotal > maxCellWidthTotal){
                    merged ++;
                    TableContent tableContentMerge = new TableContent();
                    tableContentMerge.setRowId(row);
                    tableContentMerge.setContent("HMERGE");
                    tableContentMerge.setTitle(title);
                    tableContentMerge.setColumnId(currentColumnNumber + merged);
                    maxCellWidthTotal = maxCellWidthTotal + widthList.get(maxCellIndex);
                    maxCellIndex ++;
                    cellNewMap.put(currentColumnNumber + merged, tableContentMerge);
                }
                tableContent.setColumnId(currentColumnNumber);

                cellNewMap.put(currentColumnNumber,tableContent);
                currentColumnNumber ++;
                currentColumnNumber = currentColumnNumber + merged;



            }

            tableInfoNew.put(row,cellNewMap);
        }

        System.out.println(JSONObject.toJSONString(tableInfoNew));
        return  tableInfoNew;

    }

    void readParagraph(HWPFDocument doc, Paragraph paragraph, HashMap<Integer, List<Integer>> grade){


        String content = wordGradeExtractUtils.addDocGrade(doc,paragraph,grade);
        System.out.println(content);

    }






}
