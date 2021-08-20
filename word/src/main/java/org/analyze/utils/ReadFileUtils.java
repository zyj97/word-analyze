package org.analyze.utils;




import org.apache.commons.io.FileUtils;

import org.apache.pdfbox.io.RandomAccessBuffer;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.extractor.WordExtractor;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.extractor.POIXMLTextExtractor;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;
import com.spire.doc.*;
import org.springframework.util.DigestUtils;

import java.io.*;
import java.util.List;

/**
 * @Author: zyj
 * @Date: 2021/3/9
 */
public class ReadFileUtils {
        /*** 根据文件类型返回文件内容
         *@paramfilepath
         *@return*@throwsIOException*/

        public static void main(String[] args) throws Exception {
                String path = "C:\\Users\\29927\\Desktop\\2019template.docx";
                String path1 = "C:\\Users\\29927\\Desktop\\2019kmkklmp.docx";
//    getContentByPath("C:\\Users\\29927\\Desktop\\1111.docx");
//        readAndWriterTest3("C:\\Users\\29927\\Desktop\\1111.docx");

//        try(FileInputStream is = new FileInputStream(path)) {
//                run(System.out, is);
//        } catch (FileNotFoundException e) {
//                e.printStackTrace();
//
//        } catch (IOException e) {
//                e.printStackTrace();
//
//        }
//                html(path);
                ma5(path,path1);
        }



        public static void ma5(String path1 ,String paath2) throws Exception {

                        InputStream fis =  new FileInputStream(path1);
                        BufferedInputStream bis = new BufferedInputStream(fis);
                        System.out.println(DigestUtils.md5DigestAsHex(bis));
                        bis.close();
                        fis.close();

                        InputStream fis1 =  new FileInputStream(paath2);
                        BufferedInputStream bis1 = new BufferedInputStream(fis1);
                        System.out.println(DigestUtils.md5DigestAsHex(bis1));
                        bis.close();
                        fis.close();


                }



        public static String getContentByPath(String filepath) throws IOException {
                String[] fileTypeArr = filepath.split("\\.");

                String fileType = fileTypeArr[fileTypeArr.length - 1];
                if ("doc".equals(fileType) || "docx".equals(fileType)) {
                        return readWord(filepath, fileType);

                } else if ("xlsx".equals(fileType) || "xls".equals(fileType)) {
                        return readExcel(fileType, filepath);

                } else if ("txt".equals(fileType)) {
                        return readTxt(filepath);

                } else if ("pdf".equals(fileType)) {
                        return readPdf(filepath);

                } else if ("ppt".equals(fileType) || "pptx".equals(fileType)) {
//                        return readPPT(fileType, filepath);

                } else {
                        System.out.println("不支持的文件类型！");

                }
                return "";

        }

        /*** 读取PDF中的内容
         *@paramfilePath
         *@return

         */

        public static String readPdf(String filePath) {
                FileInputStream fileInputStream = null;

                PDDocument pdDocument = null;

                String content = "";
                try {//创建输入流对象

                        fileInputStream = new FileInputStream(filePath);//创建解析器对象

                        PDFParser pdfParser = new PDFParser(new RandomAccessBuffer(fileInputStream));

                        pdfParser.parse();//pdf文档

                        pdDocument = pdfParser.getPDDocument();//pdf文本操作对象,使用该对象可以获取所读取pdf的一些信息

                        PDFTextStripper pdfTextStripper = new PDFTextStripper();

                        content = pdfTextStripper.getText(pdDocument);

                } catch (IOException e) {
                        e.printStackTrace();

                } finally {
                        try {//PDDocument对象时使用完后必须要关闭

                                if (null != pdDocument) {
                                        pdDocument.close();

                                }
                                if (null != fileInputStream) {
                                        fileInputStream.close();

                                }

                        } catch (IOException e) {
                                e.printStackTrace();

                        }

                }
                return content;

        }

        /*** 读取Excel中的内容
         *@paramfilePath
         *@return*@throwsIOException*/

        private static String readTxt(String filePath) throws IOException {
                File f = new File(filePath);
                return FileUtils.readFileToString(f, "GBK");

        }

        /*** 读取Excel中的内容
         *@paramfilePath
         *@return

         */

        private static String readExcel(String fileType, String filePath) {
                try {
                        File excel = new File(filePath);
                        if (excel.isFile() && excel.exists()) { //判断文件是否存在

                                Workbook wb;//根据文件后缀(xls/xlsx)进行判断

                                if ("xls".equals(fileType)) {
                                        FileInputStream fis = new FileInputStream(excel); //文件流对象

                                        wb = new HSSFWorkbook(fis);

                                } else if ("xlsx".equals(fileType)) {
                                        wb = new XSSFWorkbook(excel);

                                } else {
                                        System.out.println("文件类型错误!");
                                        return "";

                                }//开始解析,获取页签数

                                StringBuffer sb = new StringBuffer("");
//        for(int i=0;i<100;i++){
//
//        }

                                Sheet sheet = wb.getSheetAt(1); //读取sheet

                                sb.append(sheet.getSheetName() + "_");
                                int firstRowIndex = sheet.getFirstRowNum() + 1; //第一行是列名，所以不读

                                int lastRowIndex = sheet.getLastRowNum();
                                for (int rIndex = firstRowIndex; rIndex <= lastRowIndex; rIndex++) { //遍历行

                                        Row row = sheet.getRow(rIndex);
                                        if (row != null) {
                                                int firstCellIndex = row.getFirstCellNum();
                                                int lastCellIndex = row.getLastCellNum();
                                                for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) { //遍历列

                                                        Cell cell = row.getCell(cIndex);
                                                        if (cell != null) {
                                                                sb.append(cell.toString());

                                                        }

                                                }

                                        }


                                }
                                return sb.toString();
                        } else {
                                System.out.println("找不到指定的文件");

                        }

                } catch (Exception e) {
                        e.printStackTrace();

                }
                return "";

        }

        /*** 读取word中的内容
         *@parampath
         *@paramfileType
         *@return

         */

        public static String readWord(String path, String fileType) {
                String buffer = "";
                try (InputStream is = new FileInputStream(new File(path))) {
                        if ("doc".equals(fileType)) {


                                WordExtractor ex = new WordExtractor(is);

                                buffer = ex.getText();


                                ex.close();

                        } else if ("docx".equals(fileType)) {
                                OPCPackage opcPackage = POIXMLDocument.openPackage(path);
                                System.out.println(opcPackage);

                                POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
                                FileInputStream fis = new FileInputStream(path);
                                XWPFDocument document = new XWPFDocument(fis);
                                List<PackagePart> list = document.getAllEmbeddedParts();


                                buffer = extractor.getText();
                                System.out.println(buffer);

                                extractor.close();

                        } else {
                                System.out.println("此文件不是word文件！");

                        }

                } catch (Exception e) {
                        e.printStackTrace();

                }
                return buffer;

        }

//        private static String readPPT(String fileType, String filePath) {
//                try {
//                        if ("ppt".equals(fileType)) {
//                                PowerPointExtractor extractor = new PowerPointExtractor(new FileInputStream(new File(filePath)));
//                                return extractor.getText();
//
//                        } else if ("pptx".equals(fileType)) {
//                                return new XSLFPowerPointExtractor(POIXMLDocument.openPackage(filePath)).getText();
//
//                        }
//
//                } catch (IOException e) {
//                        e.fillInStackTrace();
//
//                } catch (XmlException e) {
//                        e.getMessage();
//
//                } catch (OpenXML4JException e) {
//                        e.getMessage();
//
//                }
//                return "";
//
//        }

//        public void decompressToUUIDDirectory(File compressFile, String baseDirectory, List<String> decompressSuffs) throws Exception {
//                List<AttachFile> attachFileList = new ArrayList<>();
//
//                //验证压缩文件
//                boolean isFile = compressFile.isFile();
//                if (!isFile){
//                        System.out.println(String.format("compressFile非文件格式！",compressFile.getName()));
//
//                }
////                String compressFileSuff = FileUtil.getFileSuffix(compressFile.getName()).toLowerCase();
////                if (!compressFileSuff.equals("zip")){
////                        System.out.println(String.format("[%s]文件非zip类型的压缩文件！",compressFile.getName()));
////                        return null;
////                }
//
//                //region 解压缩文件(zip)
//                ZipFile zip = new ZipFile(new File(compressFile.getAbsolutePath()), Charset.forName("GBK"));//解决中文文件夹乱码
//                for (Enumeration<? extends ZipEntry> entries = zip.entries(); entries.hasMoreElements();){
//                        ZipEntry entry = entries.nextElement();
//                        String zipEntryName = entry.getName();
//                        //过滤非指定后缀文件
//                        String suff = FileUtil.getFileSuffix(zipEntryName).toLowerCase();
//                        if (decompressSuffs != null && decompressSuffs.size() > 0){
//                                if (decompressSuffs.stream().filter(s->s.equals(suff)).collect(Collectors.toList()).size() <= 0){
//                                        continue;
//                                }
//                        }
//                        //创建解压目录(如果复制的代码，这里会报错，没有StrUtil，这里就是创建了一个目录来存储提取的文件，你可以换其他方式来创建目录)
//                        String groupId = StrUtil.getUUID();
//                        File group = new File(baseDirectory + groupId);
//                        if(!group.exists()){
//                                group.mkdirs();
//                        }
//                        //解压文件到目录
//                        String outPath = (baseDirectory + groupId + File.separator + zipEntryName).replaceAll("\\*", "/");
//                        InputStream in = zip.getInputStream(entry);
//                        FileOutputStream out = new FileOutputStream(outPath);
//                        byte[] buf1 = new byte[1024];
//                        int len;
//                        while ((len = in.read(buf1)) > 0) {
//                                out.write(buf1, 0, len);
//                        }
//                        in.close();
//                        out.close();
//                }
//                //endregion
//        }

        private static void html(String path) {
                Document doc = new Document(path);
                doc.saveToFile("C:\\Users\\29927\\Desktop\\b\\14.txt", FileFormat.Xml);

        }


}
