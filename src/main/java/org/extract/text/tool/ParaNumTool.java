package org.extract.text.tool;

/**
 * @ClassName ParaNumTool
 * @Description TODO
 * @Author WANGHAN756
 * @Date 2021/7/13 12:22
 * @Version 1.0
 **/
public class ParaNumTool {
    private final static char[] cs = "零一二三四五六七八九".toCharArray();
    //支持1到9999的数字
    public static String numberToChineseCount(Integer number){
        if(number==null||number<1||number>9999){
            return null;
        }
        String temp="";
        int count = 0;
        while(number>0){
            if(count==0){
                temp = cs[number%10] + "";
            }else if(count==1){
                temp = cs[number%10] + "十" + temp;
            }else if(count==2){
                temp = cs[number%10] + "佰" + temp;
            }else if(count==3){
                temp = cs[number%10] + "千" + temp;
            }
            count++;
            number/=10;
        }
        return temp;
    }

    public static String numToLetter(Integer number) {
        if(number==null){
            return "";
        }
        String str = "";
        for (byte b : (number + "").getBytes()) {
            str += (char) (b + 48);
        }
        return str;
    }

    public static void main(String[] args) {
        String s = numberToChineseCount(2382);
        System.out.println(s);
    }
}
