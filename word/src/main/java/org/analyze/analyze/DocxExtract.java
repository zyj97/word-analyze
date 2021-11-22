package org.analyze.analyze;

//import com.microsoft.schemas.vml.CTImageData;
import com.microsoft.schemas.vml.CTShape;
import org.analyze.utils.RomanUtils;
import org.analyze.utils.StringsUtils;
import org.analyze.utils.WordGradeExtractUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.json.simple.JSONArray;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

/**
 * @Author: zyj
 * @Date: 2021/6/7
 */
@Component

public class DocxExtract {
    @Autowired
    StringsUtils stringsUtils;
    @Autowired
    WordGradeExtractUtils wordGradeExtractUtils;



    public void extract(MultipartFile file) {
        List<String> imageBundleList = new ArrayList<String>();
        try (InputStream inputStream = file.getInputStream()) {

            XWPFDocument doc = new XWPFDocument(inputStream);
            System.out.println("获取属性");
            List<XWPFPictureData> picList = doc.getAllPictures();
            for (XWPFPictureData pic : picList) {
                System.out.println(pic.getFileName());
                byte[] bytev = pic.getData();
                FileOutputStream fos = new FileOutputStream(pic.getFileName());
                fos.write(bytev);
            }


            POIXMLProperties.CoreProperties properties = doc.getProperties().getCoreProperties();
            List<IBodyElement> lstElement = doc.getBodyElements();

            for (int i = 0; i < lstElement.size(); i++) {
                if (BodyElementType.PARAGRAPH.equals(lstElement.get(i).getElementType())) {
                    XWPFParagraph paragraph = (XWPFParagraph) lstElement.get(i);
                    analyzeParagraph(paragraph, doc, imageBundleList);

                } else if (BodyElementType.TABLE.equals(lstElement.get(i).getElementType())) {
                    XWPFTable table = (XWPFTable) lstElement.get(i);
                    analyzeTable(table);
                } else if (BodyElementType.CONTENTCONTROL.equals(lstElement.get(i).getElementType())) {
                    //内容控件 有边框的
                    System.out.println("不知道是什么");
                }
            }
            for (String pictureId : imageBundleList) {
                XWPFPictureData pictureData = doc.getPictureDataByID(pictureId);

                String imageName = pictureData.getFileName();
                System.out.println(imageName);
//            String lastParagraphText = paragraphList.get(i-1).getParagraphText();
//                System.out.println(pictureId +"\t|" + imageName + "\t|" + lastParagraphText);

//                System.out.println(imageName);
                File file1 = new File(imageName);

            }
        } catch (IOException e) {
//            e.printStackTrace();
        }


        JSONArray.toJSONString(imageBundleList);
    }

    private void analyzeParagraph(XWPFParagraph paragraph, XWPFDocument doc, List<String> imageBundleList) {

//        HashMap<BigInteger, List<Integer>> grade = new HashMap<>();
//        String a1 = wordGradeExtractUtils.addDocxGrade(paragraph,grade);
        String a = paragraph.getText();
//        if (StringUtils.isEmpty(a)) {
//            return;
//        }
        a = a.replaceAll(" ", "");
        //decimal
        //

//        if (StringUtils.isEmpty(a)) {
//            return;
//        }


        BigInteger numId = paragraph.getNumID();
        HashMap<BigInteger,List<Integer>> grade = new HashMap<>();
        List<Integer> numbers = grade.get(numId);
        if (numbers == null){
            numbers = new ArrayList<>();
        }
        int number = numbers.size() +1;
        number ++;
        numbers.add(number);


        String numLevelText = paragraph.getNumLevelText();

        String numFmt = paragraph.getNumFmt();
        System.out.println(numFmt);
        String afterReplacement = "";
        if (StringUtils.isNotEmpty(numFmt)){
            switch (numFmt){
                case "decimal":
                    //数字
                    afterReplacement = String.valueOf(number);
                    break;
                case  "chineseCountingThousand":
//                    汉字
                    afterReplacement = stringsUtils.cwchange(String.valueOf(number));
                    break;
                case "lowerLetter":
                    char lowerLetter = (char)(number + 96);
                    afterReplacement = String.valueOf(lowerLetter);
                    break;
                case "lowerRoman":
                    afterReplacement = RomanUtils.convert(String.valueOf(number),"lowerRoman");
                    break;
                case "upperRoman":
                    afterReplacement = RomanUtils.convert(String.valueOf(number),"lowerRoman");
                    break;
                case "upperLetter":
                    char upperLetter = (char)(number + 64);
                    afterReplacement = String.valueOf(upperLetter);
                    break;

            }
            numLevelText = numLevelText.replaceFirst("%1",afterReplacement);
            a = numLevelText.concat(a);
        }







        System.out.println(a);


        List<XWPFRun> runs = paragraph.getRuns();
//        String styleName = getDocxStyleName(doc, paragraph);
        String font = null;
        float fontSize = 0.0f;
        getSaveImgPath(doc,runs ,new Date(),"111");
        for (XWPFRun xwpfRun : runs) {






//        System.out.println("===============");
//        System.out.println(a);
//        System.out.println(styleName);

        }
    }

    public List<String> getSaveImgPath(XWPFDocument document, List<XWPFRun> runs, Date now, String fileName){
        List<String> saveImgPathList = new ArrayList<>();
        try{
            for (XWPFRun xwpfRun : runs) {
                //XWPFRun是POI对xml元素解析后生成的自己的属性，无法通过xml解析，需要先转化成CTR
                CTR ctr = xwpfRun.getCTR();

                //对子元素进行遍历
                XmlCursor c = ctr.newCursor();
                //这个就是拿到所有的子元素：
                c.selectPath("./*");
                while (c.toNextSelection()) {
                    XmlObject o = c.getObject();
                    //如果子元素是<w:drawing>这样的形式，使用CTDrawing保存图片
                    if (o instanceof CTDrawing) {
                        System.out.println("**********是图片1");
                        CTDrawing drawing = (CTDrawing) o;
                        System.out.println(drawing);
                        CTAnchor[] ctAnchors = drawing.getAnchorArray();
                        CTInline[] ctInlines = drawing.getInlineArray();
                        if (ctAnchors.length != 0) {
                            for (CTAnchor ctAnchor : ctAnchors) {
                                CTGraphicalObject graphic = ctAnchor.getGraphic();
                                XmlCursor cursor = graphic.getGraphicData().newCursor();
                                cursor.selectPath("./*");
                                while (cursor.toNextSelection()) {
                                    XmlObject xmlObject = cursor.getObject();
                                    // 如果子元素是<pic:pic>这样的形式
                                    if (xmlObject instanceof CTPicture) {
                                        CTPicture picture = (CTPicture) xmlObject;
                                        //拿到元素的属性
                                        String pictureId = picture.getBlipFill().getBlip().getEmbed();
                                        XWPFPictureData pictureData = document.getPictureDataByID(pictureId);
                                        String pictureName = pictureData.getFileName();
                                        System.out.println("===========================================格式1=====");
                                        System.out.println(pictureName);
//                                        String pictureSavePath = saveImgPath(fileName,pictureName,now);
//                                        saveImgPathList.add(pictureSavePath);
                                    }
                                }
                            }
                        }
                        if (ctInlines.length != 0) {
                            for (CTInline ctInline : ctInlines) {
                                CTGraphicalObject graphic = ctInline.getGraphic();
                                XmlCursor cursor = graphic.getGraphicData().newCursor();
                                cursor.selectPath("./*");
                                while (cursor.toNextSelection()) {
                                    XmlObject xmlObject = cursor.getObject();
                                    // 如果子元素是<pic:pic>这样的形式
                                    if (xmlObject instanceof CTPicture) {
                                        CTPicture picture = (CTPicture) xmlObject;
                                        //拿到元素的属性
//                                    XWPFPictureData pictureData = doc.getPictureDataByID(pictureId);
//
//                                    String imageName = pictureData.getFileName();
//                                    picture.
                                        String pictureId = picture.getBlipFill().getBlip().getEmbed();
                                        XWPFPictureData pictureData = document.getPictureDataByID(pictureId);
                                        String pictureName = pictureData.getFileName();
                                        System.out.println("===========================================格式2=====");
                                        System.out.println(pictureName);
                                    }
                                }
                            }
                        }
                    }
                    //使用CTObject保存图片
                    //<w:object>形式
                    if (o instanceof CTObject) {
                        System.out.println("**********是图片2");
                        System.out.println(o);
                        CTObject object = (CTObject) o;
                        System.out.println(object);
                        XmlCursor w = object.newCursor();
                        w.selectPath("./*");
                        while (w.toNextSelection()) {
                            XmlObject xmlObject = w.getObject();
                            if (xmlObject instanceof CTShape) {
                                CTShape shape = (CTShape) xmlObject;
//                                List<CTImageData> ctImageDatas = shape.getImagedataList();
//                                for (CTImageData image : ctImageDatas) {
//                                    String pictureId = image.getId2();
//                                    XWPFPictureData pictureData = document.getPictureDataByID(pictureId);
//                                    String pictureName = pictureData.getFileName();
//                                    System.out.println("===========================================格式3=====");
//                                    System.out.println(pictureName);
//                                }
                            }
                        }
                    }
                }
            }
        }catch (Exception e){
//            System.out.println("=======================================================有异常" );
            e.printStackTrace();
        }

        return saveImgPathList;

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

    private void analyzeTable(XWPFTable table) {


        List<XWPFTableRow> rows = table.getRows();
        for (int row = 0; row < rows.size(); row++) {
            XWPFTableRow xwpfTableRow = rows.get(row);
            // 获取表格的每个单元格
            List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
            for (int column = 0; column < tableCells.size(); column++) {
                XWPFTableCell cell = tableCells.get(column);
                // 获取单元格的内容
//                String text1 = cell.getText();
//                System.out.println("=========是表格====" + row+"===============" + column);
//                System.out.println(text1);

            }
        }

    }

//    private void analyzePicture(){
//        List<String> imageBundleList = XWPFUtils.readImageInParagraph(paragraphList.get(i));
//        if(CollectionUtils.isNotEmpty(imageBundleList)){
//            for(String pictureId:imageBundleList){
//                XWPFPictureData pictureData = xwpfDocument.getPictureDataByID(pictureId);
//                String imageName = pictureData.getFileName();
//                String lastParagraphText = paragraphList.get(i+1).getParagraphText();
//                System.out.println(pictureId +"\t|" + imageName + "\t|" + lastParagraphText);
//            }
//        }
//
//
//    }

}
