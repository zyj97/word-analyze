package org.analyze.utils;

import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Author: zyj
 * @Date: 2021/8/5
 */
@Component
public class StringsUtils {

    /**
     * 数字变换用数组（变换前）
     */
    final String[] numlist = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "."};

    /**
     * 数字变换用数组（变换后）
     */
    final String[] chnumlist = {"一", "二", "三", "四", "五", "六", "七", "八", "九", "零", "负", "点"};

    /**
     * 路径变换用数组
     */
    final String[] file_path_invalid_char = {"\\", "/", ":", "*", "?", "\"", ">", "<", "|"};

    private Map<Integer, String> cdict = null;
    private Map<Character, String> gdict = null;
    private Map<Integer, String> xdict = null;

    public String replacePathInvalid(String str) {
        String result = str;
        // \ / : * ? " < > |
        for (String strTmp : file_path_invalid_char) {
            result = result.replace(strTmp, " ");
        }
        return result;
    }

    /**
     * Converts numbers to Chinese representations.
     * `big`   : use financial characters.
     * `simp`  : use simplified characters instead of traditional characters.
     * `o`     : use 〇 for zero.
     * `twoalt`: use 两/兩 for two when appropriate.
     * Note that `o` and `twoalt` is ignored when `big` is used,
     * and `twoalt` is ignored when `o` is used for formal representations.
     *
     * @param num 待变换数字
     * @return 变换后数字
     * @throws Exception
     */
    public String num2chinese(String num) throws Exception {
        boolean big = false;
        boolean simp = true;
        boolean o = false;
        // check num first
        String nd = num;
        if (Double.parseDouble(nd) >= 1e48) {
            throw new Exception("number out of range");
        } else if (nd.indexOf("e") >= 0) {
            throw new Exception("scientific notation is not supported");
        }

        String[] c_symbol = simp?"正,负,点".split(","):"正,負,點".split(",");
        String[] c_basic = null;
        String[] c_unit1 = null;
        String c_twoalt = null;
        if (big) {
            c_basic = simp?"零,壹,贰,叁,肆,伍,陆,柒,捌,玖".split(","):"零,壹,贰,叁,肆,伍,陆,柒,捌,玖".split(",");
            c_unit1 = "拾,佰,仟".split(",");
            c_twoalt = simp?"贰":"貳";
        } else {
            c_basic = simp?"零,壹,贰,叁,肆,伍,六,柒,捌,玖".split(","):"零,壹,贰,叁,肆,伍,六,柒,捌,玖".split(",");
            c_unit1 = "拾,佰,仟".split(",");
            c_twoalt = simp?"贰":"貳";
        }

        String[] c_unit2 = simp?"万,亿,兆,京,垓,秭,穰,沟,涧,正,载".split(","):"万,亿,兆,京,垓,秭,穰,沟,涧,正,载".split(",");

        String result = "";
        int index = 0;
        if (nd.charAt(0) == '+') {
            result = result + c_symbol[0];
            index = 1;
        } else if (nd.charAt(0) == '-') {
            result = result + c_symbol[1];
            index = 1;
        }

        String tmpStr = nd.substring(index);
        String integer = tmpStr;
        String remainder = "";
        if (nd.indexOf(".") >= 0) {
            integer = tmpStr.split("\\.")[0];
            remainder = tmpStr.split("\\.")[1];
        }

        if (isNumeric(integer)) {
            String spl = "";
            List<String> splitted = new ArrayList<>();
            for (int i = integer.length() - 1; i >= 0; i--) {
                spl = integer.charAt(i) + spl;
                if (i == 0) {
                    splitted.add(spl);
                } else if (spl.length() == 4) {
                    splitted.add(spl);
                    spl = new String("");
                }
            }
            String intresult = "";

            for (int i = 0; i < splitted.size(); i++) {
                String unit = splitted.get(i);
                if (Integer.parseInt(unit) == 0) {
                    intresult = intresult + c_basic[0];
                    continue;
                } else if (i > 0 && Integer.parseInt(unit) == 2) {
                    intresult = intresult + c_twoalt + c_unit2[i - 1];
                    continue;
                }
                String ulist = "";
                unit = String.format("%04d", Integer.parseInt(unit));
                for (int j = unit.length() - 1; j >= 0; j--) {
                    char ch = unit.charAt(j);
                    if (ch == '0' && !"0".equals(ulist)) {
                        ulist = c_basic[0] + ulist;
                    } else if (j == unit.length() - 1) {
                        ulist = c_basic[Character.getNumericValue(ch)] + ulist;
                    } else if (j == unit.length() - 2 && ch == '1' && unit.charAt(1) == '0') {
                        ulist = c_unit1[0] + ulist;
                    } else if (j < unit.length() - 2 && ch == '2') {
                        ulist = c_twoalt + c_unit1[unit.length() - 2 - j] + ulist;
                    } else {
                        ulist = c_basic[Character.getNumericValue(ch)] + c_unit1[unit.length() - 2 - j] + ulist;
                    }
                }
                if (i == 0) {
                    intresult = ulist + intresult;
                } else {
                    intresult = ulist + c_unit2[i - 1] + intresult;
                }
            }
            result = trimChars(intresult, c_basic[0]);
        } else {
            result = c_basic[0];
        }

        if (!remainder.isEmpty() && isNumeric(remainder) && Integer.parseInt(remainder) != 0) {
            result = result + c_symbol[2];
            for (int i = 0; i < remainder.length(); i++) {
                char ch = remainder.charAt(i);
                result = result + c_basic[Character.getNumericValue(ch)];
            }
        }

        return result;
    }

    public String text2allchinese(String text, int style) {
        // 1: n
        // 2: t
        if (style == 1) {
            String result = "";
            List<String> lst = new ArrayList<>();
            boolean flg = false; // 数字:TRUE 以外:FALSE
            int idxStart = 0;
            for (int i = 0; i < text.length(); i++) {
                if (i == 0) {
                    if (Character.isDigit(text.charAt(i))) {
                        flg = true;
                    } else {
                        flg = false;
                    }
                } else {
                    if (Character.isDigit(text.charAt(i)) && !flg) {
                        lst.add(text.substring(idxStart, i));
                        flg = true;
                        idxStart = i;
                    } else if (!Character.isDigit(text.charAt(i)) && flg) {
                        lst.add(text.substring(idxStart, i));
                        flg = false;
                        idxStart = i;
                    }
                    if (i == text.length() - 1) {
                        lst.add(text.substring(idxStart));
                    }
                }
            }
            boolean flagPoint = false;
            for (int j = 0; j < lst.size(); j++) {
                try {
                    if (lst.get(j).indexOf(".") >= 0) {
                        flagPoint = true;
                    }
                    if (!flagPoint) {
                        result = result + num2chinese(lst.get(j));
                    } else {
                        String strRep = lst.get(j);
                        for (int k = 0; k < numlist.length; k++) {
                            String tempStr = strRep;
                            strRep = tempStr.replace(numlist[k], chnumlist[k]);
                        }
                        result = result + strRep;
                    }
                } catch (Exception e) {
                    result = result + lst.get(j);
                }
            }
            return result;
        } else {
            String strRep = text;
            for (int k = 0; k < numlist.length; k++) {
                String tempStr = strRep;
                strRep = tempStr.replace(numlist[k], chnumlist[k]);
            }
            return strRep;
        }
    }

    /*
     * 拆分函数，将整数字符串拆分成[亿，万，仟]的list
     */
    private List<String> csplit(String cdata) {
        int g = cdata.length() % 4;
        List<String> csdata = new ArrayList<>();
        int lx = cdata.length() - 1;
        if (g > 0) {
            csdata.add(cdata.substring(0, g));
        }
        int k = g;
        while (k <= lx) {
            csdata.add(cdata.substring(k, k + 4));
            k += 4;
        }
        return csdata;
    }

    /*
     * 对[亿，万，仟]的list中每个字符串分组进行大写化再合并
     */
    private String cschange(String cki) {
        int lenki = cki.length();
        int lk = lenki;
        String chk = "";
        for (int i = 0; i < lenki; i++) {
            if (cki.charAt(i) == '0') {
                if (i < lenki - 1) {
                    if (cki.charAt(i + 1) != '0') {
                        chk = chk + this.getGdict().get(cki.charAt(i));
                    }
                }
            } else {
                chk = chk + this.getGdict().get(cki.charAt(i)) + this.getCdict().get(lk);
            }
            lk -= 1;
        }
        return chk;
    }

    public String cwchange(String data) {
        if (data.indexOf(".") < 0) {
            data = data + ".0";
        }

        String[] cdata = data.split("\\.");
        String cki = cdata[0];
        String ckj = cdata[1];
        String chk = "";

        // 分解字符数组[亿，万，仟]三组List:['0000','0000','0000']
        List<String> cski = this.csplit(cki);
        // 获取拆分后的List长度
        int ikl = cski.size();

        // 大写合并
        for (int i = 0; i < ikl; i++) {
            // 有可能一个字符串全是0的情况
            if ("".equals(this.cschange(cski.get(i)))) {
                // 此时不需要将数字标识符引入
                chk = chk + this.cschange(cski.get(i));
            } else {
                // 合并：前字符串大写+当前字符串大写+标识符
                chk = chk + this.cschange(cski.get(i));
            }
        }

        // 处理小数部分
//        int lenkj = ckj.length();
//        // 若小数只有1位
//        if (lenkj == 1) {
//            if (ckj.charAt(0) == '0') {
//                chk = chk + "整";
//            } else {
//                chk = chk + this.getGdict().get(cki.charAt(0)) + "角整";
//            }
//        } else {
//            if (ckj.charAt(0) == '0' && ckj.charAt(1) == '0') {
//                chk = chk + "整";
//            } else if (ckj.charAt(0) == '0' && ckj.charAt(1) != '0') {
//                chk = chk + "零" + this.getGdict().get(cki.charAt(1)) + "分";
//            } else if (ckj.charAt(0) != '0' && ckj.charAt(1) != '0') {
//                chk = chk + this.getGdict().get(cki.charAt(0)) + "角" + this.getGdict().get(cki.charAt(1)) + "分";
//            } else {
//                chk = chk + this.getGdict().get(cki.charAt(0)) + "角整";
//            }
//        }
        return chk;
    }

    public static boolean isNumeric(String str) {
        if (str == null) {
            return false;
        }
        for (int i = 0; i < str.length(); i++) {
            if (!Character.isDigit(str.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    /*
     * 去除字符串两端的指定字符
     */
    public static String trimChars(String str, String trimStr){
        for(int i = 0; i < str.length(); i++){
            // 从字符串的第一个字符往后遍历
            if(trimStr.equals(String.valueOf(str.charAt(i)))){
                // 遍历到指定字符就用空白字符替换，直到遍历到不是指定字符为止
                str = str.substring(0, i + 1).replace(str.charAt(i), ' ') + str.substring(i + 1);
            }else{
                for(int j = str.length() - 1; j >= 0; j--){
                    // 从字符串的最后一个字符往前遍历
                    if(trimStr.equals(String.valueOf(str.charAt(j)))){
                        // 遍历到指定字符就用空白字符替换，直到遍历到不是指定字符为止
                        str = str.substring(0, j) + str.substring(j).replace(str.charAt(j), ' ');
                    }else{
                        break;
                    }
                }
                break;
            }
        }
        // 去除字符串两端的空白字符并返回
        return str.trim();
    }

    private Map<Integer, String> getCdict() {
        cdict = new HashMap<>();
        cdict.put(1, "");
        cdict.put(2, "十");
        cdict.put(3, "百");
        cdict.put(4, "千");
        return cdict;
    }

    private Map<Character, String> getGdict() {
        gdict = new HashMap<>();
        gdict.put('0', "零");
        gdict.put('1', "一");
        gdict.put('2', "二");
        gdict.put('3', "三");
        gdict.put('4', "四");
        gdict.put('5', "五");
        gdict.put('6', "六");
        gdict.put('7', "七");
        gdict.put('8', "八");
        gdict.put('9', "九");
        return gdict;
    }

    private Map<Integer, String> getXdict() {
        xdict = new HashMap<>();
        xdict.put(1, "元");
        xdict.put(2, "万");
        xdict.put(3, "亿");
        xdict.put(4, "兆");
        return xdict;
    }
}
