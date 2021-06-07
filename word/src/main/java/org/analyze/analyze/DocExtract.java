package org.analyze.analyze;

import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;

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
        InputStream inputStream = null;
        try{
            inputStream = file.getInputStream();
            POIFSFileSystem pfs = new POIFSFileSystem(inputStream);
            HWPFDocument hwpf = new HWPFDocument(pfs);

            Range range = hwpf.getRange();//得到文档的读取范围
            TableIterator it = new TableIterator(range);
            for (int o = 0; o< range.numCharacterRuns();o++){
                short i = range.getCharacterRun(o).getStyleIndex();
                System.out.println("============格式为" + i);
                System.out.println(range.getCharacterRun(o).getFontName());
                System.out.println(range.getCharacterRun(o).getFontSize());
                System.out.println("颜色为" +range.getCharacterRun(o).getColor());

                System.out.println("内容" +range.getCharacterRun(o).getParagraph(0).text());
            }


            if (it.hasNext()) {
                Table table = it.next();
                for (int row = 0; row < table.numRows(); row++) {
                    TableRow rows = table.getRow(row);

                    for (int column = 0; column < rows.numCells(); column++) {

                        Paragraph f = rows.getParagraph(column);
                        TableCell tableCell = rows.getCell(column);
                        System.out.println("==================row");

                        System.out.println("第" + (row + 1) + "行" + "第" + (column + 1) + "列" + "填写内容为： " + tableCell.text() + "======");

                    }
                }
            }



        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            inputStream.close();
        }


    }





}
