package org.analyze.utils;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.ListLevel;
import org.apache.poi.hwpf.model.ListTables;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.stereotype.Component;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * @Author: zyj
 * @Date: 2021/9/6
 */
@Component
public class WordGradeExtractUtils {



    /**
     * 添加doc格式文件的自动标号
     *
     * @param doc
     * @param paragraph
     * @param grade
     * @return
     */
    /**
     *添加自动标号
     */
    public String addDocGrade(HWPFDocument doc,Paragraph paragraph,HashMap<Integer, List<Integer>> grade){

        String content = paragraph.text();
        if (!paragraph.isInList()) {
            return content;
        }
        ListTables lTable = doc.getListTables();
        if (lTable == null){
            return content;
        }
        ListLevel ll = lTable.getLevel(paragraph.getList().getLsid(), paragraph.getIlvl());

        int numId = paragraph.getList().getLsid();

        List<Integer> numbers = grade.get(numId);
        if (numbers == null){
            numbers = new ArrayList<>();
        }
        int number = numbers.size() +1;
        numbers.add(number);
        grade.put(numId,numbers);


        String numLevelText = ll.getNumberText();

        int numFmt = ll.getNumberFormat();
//        System.out.println(numFmt);
        String afterReplacement = "";

        switch (numFmt){
            case 0:
                //数字
                afterReplacement = String.valueOf(number);
                break;
            case 11:
//                    日语
                afterReplacement = ChineseNumberUtils.cwchange(String.valueOf(number));
                break;
            case  39:
//                    汉字
                afterReplacement = ChineseNumberUtils.cwchange(String.valueOf(number));
                break;
            case 4:
                //小写字母
                char lowerLetter = (char)(number + 96);
                afterReplacement = String.valueOf(lowerLetter);
                break;
            case 2:
                //小写罗马字母
                afterReplacement = RomanUtils.convert(String.valueOf(number),"lowerRoman");
                break;
//                case :
//                大写罗马字母 遇到了再加
//                    afterReplacement = RomanUtils.convert(String.valueOf(number),"lowerRoman");
//                    break;
            case 3:
                //大写字母
                char upperLetter = (char)(number + 64);
                afterReplacement = String.valueOf(upperLetter);
                break;

        }
        if (numLevelText.startsWith("(") || numLevelText.startsWith("（")){
            StringBuffer sb = new StringBuffer(numLevelText);
            sb.insert(1,afterReplacement);
            numLevelText = sb.toString();
        }else {
            numLevelText = afterReplacement.concat(numLevelText);
        }

        content = numLevelText.concat(content);
//        System.out.println(content);



        return content;
    }

    /**
     * 添加docx格式文件的自动标号
     * @param paragraph
     * @param grade
     * @return
     */
    public  String addDocxGrade(XWPFParagraph paragraph, HashMap<BigInteger, List<Integer>> grade){
        String content = "";
        BigInteger numId = paragraph.getNumID();
        if (null == numId){
            return content;
        }

        List<Integer> numbers = grade.get(numId);
        if (numbers == null){
            numbers = new ArrayList<>();
        }
        int number = numbers.size() +1 ;
        numbers.add(number);
        grade.put(numId,numbers);

        String numLevelText = paragraph.getNumLevelText();
        String numFmt = paragraph.getNumFmt();
//        System.out.println(numFmt);
        String afterReplacement = "";
        if (StringUtils.isNotEmpty(numFmt)){
            switch (numFmt){
                case "decimal":
                    //数字
                    afterReplacement = String.valueOf(number);
                    break;
                case  "chineseCountingThousand":
//                    汉字
                    afterReplacement = ChineseNumberUtils.cwchange(String.valueOf(number));
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
                case "japaneseCounting":
                    afterReplacement = ChineseNumberUtils.cwchange(String.valueOf(number));
                    break;
                default:
                    afterReplacement = String.valueOf(number);
                    break;

            }
            numLevelText = numLevelText.replaceFirst("%1",afterReplacement);
            content = numLevelText.concat(content);

        }
        return content;
    }
}
