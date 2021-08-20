package org.analyze.service;


import org.analyze.analyze.XWPFWordExtract;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

/**
 * @Author: zyj
 * @Date: 2021/8/20
 *
 */
@Service
public class GenerateMdFile {

    @Autowired
    XWPFWordExtract xwpfWordExtract;
    private static String datasource = "### 数据库说明\n";
    private static String datasourceSql = "### 数据库sql\n";



    public String ganderMdfile(MultipartFile file) throws Exception {
        String result = "";

        //读取文档
        String content =xwpfWordExtract.testReadByExtractor(file);

        String[] strings = content.split("create table");
        for (int i = 1;i<strings.length +1;i++){
            //生成
            String mdString =strings[i-1];
            String trim = mdString.replaceAll(" ","");
            if (trim.length()<=0){
                continue;
            }
            mdString = "create table ".concat(mdString);
            String table = ganderTable(mdString);
            String tableName = ganderTableName(mdString);


            String tableNameComment = ganderTableNameComment(mdString);
            String title = "# ".concat(i-1 +".").concat(tableNameComment.replace(";","")).concat("(").concat(tableName).concat(")")
                    .concat("\n");
            String re = title.concat(datasource).concat(table).concat("\n").concat("\n").concat(datasourceSql).
                    concat("```sql\n").concat(mdString).concat("\n").concat("```");
            result = result.concat("\n").concat(re);


        }
        return result;




    }

    private String ganderTable(String str){
        String gander = "| 顺序 | 字段名称      | 类型        | 是否可以为空 | 是否自增 | 默认值            | 说明         | 备注                             |\n" +
                "| ---- | ------------- | ----------- | ------------ | -------- | ----------------- | ------------ | -------------------------------- |";
        int first = StringUtils.indexOf(str,"(");
        int last = StringUtils.lastIndexOf(str,")");
        String file = StringUtils.substring(str,first +1,last -1);
        String[] list = StringUtils.split(file,",");
        String lastStr = "";

        //引号开始和结束
        int singleQuoteStart = -2;
        int singleQuoteEnd = -2;

        //括号开始和结束
        int bracketsStart = -2;
        int bracketsEnd = -2;



        for (int i= 1;i<list.length +1;i++){

            String row = list[i-1];

            int start = StringUtils.indexOf(row,"(");
            int end = StringUtils.lastIndexOf(row,")");

            int singleQuote = StringUtils.indexOf(row,"'");
            int singleQuoteLast = StringUtils.lastIndexOf(row ,"'");
            if (singleQuote >=singleQuoteLast){
                //这个是逗号在单引号中
                if (singleQuote  > -1 && singleQuoteStart < 0){
                    //出现第一个单引号
                    singleQuoteStart = singleQuote;
                    lastStr = row;
                    continue;

                }else if (singleQuote  > -1 && singleQuoteStart > 0){
                    //出现第二个单引号
                    singleQuoteEnd = singleQuote;
                }
                if (singleQuoteStart >=0 &&singleQuoteEnd <0){
                    row = row.replaceFirst(" ","");
                    lastStr = lastStr.trim();
                    lastStr = lastStr.concat(",").concat(row);
                    continue;
                } else if ((singleQuoteStart >0 && singleQuoteEnd > 0)){
                    lastStr = lastStr.concat(",").concat(row);
                    row = lastStr;
                    singleQuoteStart = -2;
                    singleQuoteEnd = -2;
                    lastStr = "";

                }
            }



            if (start <0 || end <0){
                if (start > -1){
                    bracketsStart = start;
                }
                if (end > -1){
                    bracketsEnd = end;
                }
            }

            if (bracketsStart >= 0 && bracketsEnd >=0 && StringUtils.isNotEmpty(lastStr)){
                row = row.replaceFirst(" ","");
                lastStr = lastStr.trim().concat(",").concat(row);
                row = lastStr;
                bracketsStart = -2;
                bracketsEnd = -2;
                lastStr = "";
            }else if (bracketsStart >= 0 && bracketsEnd < 0){
                if (StringUtils.isNotEmpty(lastStr)){
                    row = row.replaceFirst(" ","");
                    lastStr = lastStr.trim();
                    lastStr = lastStr.concat(",").concat(row);
                }else {
                    lastStr = row;
                }

                continue;
            }
            //这是逗号在括号中

            String a = row.replaceAll("\n","").replaceAll(" ","");
            if (a.startsWith("primarykey") || a.startsWith("constraint")){
                continue;
            }
            String[] rowList = row.split(" ");
            String name= "";
            String type = "";
            String isNull = "√";
            String isAddition = "×";
            String defaultValue = "null";
            String comment = "";
            String remark = "";
            int index = 0;
            if (StringUtils.contains(row,"not null")){
                isNull = "×";
            }
            if (StringUtils.contains(row,"auto_increment")){
                isAddition = "√";
            }
            if (StringUtils.contains(row,"default")){
                isAddition = "√";
            }

            boolean hasDefault =false;
            boolean hasComment = false;
            for (String s: rowList) {

                s = s.replaceAll(" ","");
                s = s.replaceAll("\n","");
                if (s.length() <=0){
                    continue;
                }
                s = s.replaceAll("'","");

                if (StringUtils.isEmpty(name)){
                    name = s;
                    continue;
                }else if (StringUtils.isEmpty(type)){
                    type = s;
                    continue;
                }
                if ("default".equals(s)){
                    hasDefault = true;
                }else if (hasDefault){
                    defaultValue = s;
                    hasDefault = false;
                }
                if ("comment".equals(s)){
                    hasComment = true;
                }else if (hasComment && index <=2){
                    comment = comment.concat(" ").concat(s);
                }


            }
            String s ="|".concat(i +"   ")
                    .concat("|").concat(name)
                    .concat("|").concat(type)
                    .concat("|").concat(isNull)
                    .concat("|").concat(isAddition)
                    .concat("|").concat(defaultValue)
                    .concat("|").concat(comment)
                    .concat("|").concat(remark);

            gander = gander.concat("\n").concat(s);


        }
        return gander;
    }

    /**
     * 获取数据表名称
     * @param sql
     * @return
     */
    private String ganderTableName(String sql){

        sql = sql.replace("create table ","");
        int firstIndex = StringUtils.indexOf(sql,"(");
        int lastIndex = StringUtils.lastIndexOf(sql,")");
        StringBuffer stringBuffer = new StringBuffer(sql);
        stringBuffer = stringBuffer.replace(firstIndex,lastIndex +1,"");
        String str = stringBuffer.toString();
        String[] strings = str.split(" ");

        String result = "";
        for (String a : strings){
            a = a.replaceAll("\n","");
            a = a.replaceAll(" ","");
            if (StringUtils.isEmpty(a)){
                continue;
            }
            a = a.replaceAll("","");

            result = a;
            break;

        }
        return result;
    }

    /**
     * 获取表名备注
     * @param sql
     * @return
     */
    private String  ganderTableNameComment(String sql){
        int firstIndex = StringUtils.indexOf(sql,"(");
        int lastIndex = StringUtils.lastIndexOf(sql,")");
        StringBuffer stringBuffer = new StringBuffer(sql);
        stringBuffer = stringBuffer.replace(firstIndex,lastIndex,"");
        String str = stringBuffer.toString();
        String[] strings = str.split(" ");
        Boolean hasComment = false;
        String result = "";
        int index = 0;
        for (String a : strings){
            a = a.replaceAll("\n","");
            a = a.replaceAll(" ","");
            if (StringUtils.isEmpty(a)){
                continue;
            }
            if (index >= 2){
                break;
            }
            if (StringUtils.contains(a,"'")){
                index ++;

            }

            a = a.replaceAll("'","");

            if ("comment".equals(a)){
                hasComment = true;
            }else if (hasComment && index <=2){
                result = result.concat(" ").concat(a);
            }

        }
        return result;
    }
}
