package org.extract.text.tool;

import com.google.gson.ExclusionStrategy;
import com.google.gson.FieldAttributes;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

import org.apache.commons.lang.StringEscapeUtils;

import org.extract.pojo.Tu;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @ClassName TextTool
 * @Description 涉及到文本处理的工具类
 * @Author WANGHAN756
 * @Date 2021/4/21 9:20
 * @Version 1.0
 **/
public class TextTool {

    /**
     * 替换字符串&为和，这是因为&在dom4j中属于特殊字符，会影响解析dom
     * @param str 输入字符串
     * @return 返回替换后的字符串
     */
    public static String reKeyAnd(String str){
        if(str==null){
            return "";
        }
        return str.replaceAll("&","和");
    }
    /**
     * 反向替换
     * @param str 输入字符串
     * @return 返回替换后的字符串
     */
    public static String bakKeyAnd(String str){
        if(str==null){
            return "";
        }
        return str.replaceAll("和","&");
    }

    /**
     * 返回json str
     * @param object 对象
     * @return json str
     */
    public static String toJson(Object object) {
        Gson gson = new GsonBuilder().disableHtmlEscaping().setPrettyPrinting().addSerializationExclusionStrategy(new ExclusionStrategy() {
            @Override
            public boolean shouldSkipField(FieldAttributes fa) {
                return fa.getName().equals("startLine") || fa.getName().equals("startLineStatus") ||
                        fa.getName().equals("endLine") || fa.getName().equals("endLineStatus")
                        ||fa.getName().equals("WordStyleStruct")
                        || fa.getName().equals("pdfStyleStructs")  || fa.getName().equals("trueColIndex")
                        ||fa.getName().equals("vMerge");
//                return fa.getName().equals("startLine") || fa.getName().equals("startLineStatus") ||
//                        fa.getName().equals("endLine") || fa.getName().equals("endLineStatus")
//                        || fa.getName().equals("trueColIndex")
//                        ||fa.getName().equals("vMerge");
            }

            @Override
            public boolean shouldSkipClass(Class<?> arg0) {
                // TODO Auto-generated method stub
                return false;
            }
        }).create();
        // pretty print
        // GsonBuilder gsonBuilder = new GsonBuilder();
        // gsonBuilder.setPrettyPrinting();
        // Gson gson = gsonBuilder.create();
        return gson.toJson(object);
    }

    /**
     * 将escape过的str转换回来
     * @param str 输入字符串
     * @return 输出字符串
     */
    public static String unescape(String str){
        return StringEscapeUtils.unescapeHtml(str);
    }

    /**
     * 将str被escape
     * @param str 输入字符串
     * @return 输出字符串
     */
    public static String escape(String str){
        return StringEscapeUtils.escapeHtml(str);
    }

    /**
     * 将文本及其位置信息封装入String
     * @param text 文本内容
     * @param xStart 文本x起始位置
     * @param yStart 文本y起始位置
     * @param xEnd 文本x结束位置
     * @param yEnd 文本y结束位置
     * @return 含有文本内容及位置信息的String
     */
    public static String encodeTextLine(String text,float xStart,float yStart,float xEnd,float yEnd){
        DecimalFormat decimalFormat = new DecimalFormat("000.000");
        String xStartStr = decimalFormat.format(xStart);
        String yStartStr = decimalFormat.format(yStart);
        String xEndStr = decimalFormat.format(xEnd);
        String yEndStr = decimalFormat.format(yEnd);
        return xStartStr+yStartStr+xEndStr+yEndStr+text;
    }

    /**
     * 将文本内容和位置信息拆分开来
     * @param line 文本内容及位置信息
     * @return 拆分后的数组
     */
    public static String[] decodeTextLine(String line){
        String xStart = line.substring(0, 7);
        String yStart = line.substring(7, 14);
        String xEnd = line.substring(14, 21);
        String yEnd = line.substring(21, 28);
        String text = line.substring(28, line.length());
        return new String[]{text,xStart,yStart,xEnd,yEnd};
    }


    /**
     * 是否包含中文
     * @param str 字符串
     * @return 是或不是
     */
    public static boolean isContainChinese(String str){
        String reg = "[\\u4E00-\\u9FA5]+";
        Pattern p = Pattern.compile(reg);
        Matcher m = p.matcher(str);
        if(m.find()){
            return true;
        }else{
            return false;
        }
    }
    /**
     * 是否包含英文
     * @param str 字符串
     * @return 是或不是
     */
    public static boolean isContainEnglish(String str){
        String reg = "[A-Za-z]";
        Pattern p = Pattern.compile(reg);
        Matcher m = p.matcher(str);
        if(m.find()){
            return true;
        }else{
            return false;
        }
    }

    //对无法显示在网页上的unicode做了大致总结，不保证全面
    //[888,889] [896,899] [907] [909] [1367,1368] [2043,2207] [5873,6015]  [55204,55215] [55239,55242] [55292,63743]
    //[64110,64255] [64263,64274] [64280,64284] [64450.64466] [64512,64605] [64612,64753]
    //[64757,64827] [64832,65009] [65013] [65013,65017] [65022,65023] [65050,65055]
    //[65063,65071] [65107] [65127] [65132,65135] [65141] [65277,65278] [65280].
    //[65496,65497] [65501,65503] [65511] [65519,65528] [65532,65535]
    private static final Set<Integer> specialUnicode = new HashSet<>();
    static {
        List<Tu.Tuple2<Integer,Integer>> ranges = new ArrayList<>();
        ranges.add(new Tu.Tuple2<>(888,889));ranges.add(new Tu.Tuple2<>(896,899));
        ranges.add(new Tu.Tuple2<>(907,907));ranges.add(new Tu.Tuple2<>(909,909));
        ranges.add(new Tu.Tuple2<>(1367,1368));ranges.add(new Tu.Tuple2<>(2043,2207));
        ranges.add(new Tu.Tuple2<>(5873,6015));ranges.add(new Tu.Tuple2<>(55204,55215));
        ranges.add(new Tu.Tuple2<>(55239,55242));ranges.add(new Tu.Tuple2<>(55292,63743));
        ranges.add(new Tu.Tuple2<>(64110,64255));ranges.add(new Tu.Tuple2<>(64263,64274));
        ranges.add(new Tu.Tuple2<>(64280,64284));ranges.add(new Tu.Tuple2<>(64450,64466));
        ranges.add(new Tu.Tuple2<>(64512,64605));ranges.add(new Tu.Tuple2<>(64612,64753));
        ranges.add(new Tu.Tuple2<>(64757,64827));ranges.add(new Tu.Tuple2<>(64832,65009));
        ranges.add(new Tu.Tuple2<>(65013,65013));ranges.add(new Tu.Tuple2<>(65013,65017));
        ranges.add(new Tu.Tuple2<>(65022,65023));ranges.add(new Tu.Tuple2<>(65050,65055));
        ranges.add(new Tu.Tuple2<>(65063,65071));ranges.add(new Tu.Tuple2<>(65107,65107));
        ranges.add(new Tu.Tuple2<>(65127,65127));ranges.add(new Tu.Tuple2<>(65132,65135));
        ranges.add(new Tu.Tuple2<>(65141,65141));ranges.add(new Tu.Tuple2<>(65277,65278));
        ranges.add(new Tu.Tuple2<>(65280,65280));ranges.add(new Tu.Tuple2<>(65496,65497));
        ranges.add(new Tu.Tuple2<>(65501,65503));ranges.add(new Tu.Tuple2<>(65511,65511));
        ranges.add(new Tu.Tuple2<>(65519,65528));ranges.add(new Tu.Tuple2<>(65532,65535));
        ranges.forEach(x->{
            Integer start = x.getKey();
            Integer end = x.getValue();
            for(int i=start;i<=end;i++){
                specialUnicode.add(i);
            }
        });
    }


}
