package org.analyze.utils;

/**
 * @Author: zyj
 * @Date: 2021/8/5
 */

public class RomanUtils {

    public static final String[][] upperRoman = {
            {"","I","II","III","IV","V","VI","VII","VIII","IX"},  // 个位数举例
            {"","X","XX","XXX","XL","L","LX","LXX","LXXX","XC"},  // 十位数举例
            {"","C","CC","CCC","CD","D","DC","DCC","DCCC","CM"},  // 百位数举例
            {"","M","MM","MMM"}  // 千位数举例
    };

    public static final String[][] lowerRoman = {
            {"","i","ii","iii","iv","v","vi","vii","viii","ix"},  // 个位数举例
            {"","x","xx","xxx","xl","l","lx","lxx","lxxx","xc"},  // 十位数举例
            {"","c","cc","ccc","cd","d","dc","dcc","dccc","cm"},  // 百位数举例
            {"","m","mm","mmm"}  // 千位数举例
    };





    /**
     * @param strNubmer
     * @return
     *
     * 思路是：千位先除以1000然后模10，百位先除以100然后模10，十位先除以10然后模10，个位直接模10.
     */
    public static String convert(String strNubmer,String romanStr) {
        StringBuilder strBuilder = new StringBuilder();
        int number = Integer.parseInt(strNubmer);
        if ("upperRoman".equals(romanStr)){
            strBuilder.append(upperRoman[3][number/1000%10]);
            strBuilder.append(upperRoman[2][number/100%10]);
            strBuilder.append(upperRoman[1][number/10%10]);
            strBuilder.append(upperRoman[0][number%10]);
        }else {
            strBuilder.append(lowerRoman[3][number/1000%10]);
            strBuilder.append(lowerRoman[2][number/100%10]);
            strBuilder.append(lowerRoman[1][number/10%10]);
            strBuilder.append(lowerRoman[0][number%10]);

        }


        return strBuilder.toString();
    }

}
