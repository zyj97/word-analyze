package org.analyze.analyze;

import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;
import com.spire.pdf.PdfPageBase;
import com.spire.pdf.tables.PdfTable;

import java.io.FileWriter;
import java.io.IOException;


/**
 * @Author: zyj
 * @Date: 2021/8/27
 */
public class SpirePdfExtract {
    public static void main(String[] args) {
        PdfDocument doc = new PdfDocument();
        //加载PDF文件
        doc.loadFromFile("D:\\download\\dingding\\[1800014068_签章版]_6110125074(化工销售江苏分公司_其他烷烃环氧丙烷)61fd55a1872d47d592ca0e90f89ca65b.pdf");
//        doc.saveToFile("ToHTML.html", FileFormat.HTML);
        StringBuilder sb = new StringBuilder();

        PdfPageBase page;
        //遍历PDF页面，获取每个页面的文本并添加到StringBuilder对象
        for(int i= 0;i<doc.getPages().getCount();i++){
            page = doc.getPages().get(i);
            PdfTable table = new PdfTable();
            sb.append(page.extractText(true));

        }
        FileWriter writer;
        try {
            //将StringBuilder对象中的文本写入到文本文件
            writer = new FileWriter("ExtractText.txt");
            writer.write(sb.toString());
            writer.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }

        doc.close();
    }


}
