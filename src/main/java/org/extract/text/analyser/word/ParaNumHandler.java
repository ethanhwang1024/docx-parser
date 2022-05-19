package org.extract.text.analyser.word;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.extract.text.tool.ParaNumTool;

import java.util.HashMap;
import java.util.Map;

/**
 * @ClassName ParagraphNumHandler
 * @Description word段落编号处理,为什么需要这个类呢，是因为word中的自动序号并不是硬编码
 *              在文档中的，需要结合上下文来推断当前的序号
 * GetNumFmt: decimal, GetNumID: 1, GetNumIlvl: 0, NumLevelText: %1. => 1.
 * GetNumFmt: decimal, GetNumID: 4, GetNumIlvl: 0, NumLevelText: %1) => 1)
 * GetNumFmt: chineseCountingThousand, GetNumID: 2, GetNumIlvl: 0, NumLevelText: (%1) => (一)
 * GetNumFmt: chineseCountingThousand, GetNumID: 3, GetNumIlvl: 0, NumLevelText: %1、 => 一、
 * GetNumFmt: upperLetter, GetNumID: 5, GetNumIlvl: 0, NumLevelText: %1. => A.
 * GetNumFmt: decimal, GetNumID: 6, GetNumIlvl: 0, NumLevelText: %1、 => 1、
 * @Author WANGHAN756
 * @Date 2021/6/29 13:42
 * @Version 1.0
 **/

public class ParaNumHandler {
    //记录上一次的等级，低级别会受到上一次高级别的影响，导致本级别序号重置
    private Integer preLevel;


    //Num字典
    private Map<String,Integer> map = new HashMap<>();


    public String getParagraphNum(XWPFParagraph paragraph) {
        String result = null;
        if (StringUtils.isEmpty(paragraph.getNumFmt()) ||
                StringUtils.isEmpty(paragraph.getNumID()+"") ||
                StringUtils.isEmpty(paragraph.getNumLevelText())) {
            return null;
        }

        String key = paragraph.getNumID() + ","+ paragraph.getNumFmt()+","+paragraph.getNumLevelText();
//        string key = paragraph.GetNumID() ?? "";
        int curLevel = paragraph.getNumIlvl().intValue();
        if(map.containsKey(key)){
            if(curLevel>preLevel){
                //编号重置
                map.put(key,1);
            }else{
                //编号延续
                map.put(key,map.get(key)+1);
            }
        }else{
            map.put(key,1);
        }
        preLevel = curLevel;
//        String fmt = (paragraph.getNumLevelText() + "").replaceAll("%1", "{0}");
        Integer index = map.get(key);

//        System.out.println("当前:"+paragraph.getNumFmt());
//        System.out.println(paragraph.getNumLevelText());
        if(paragraph.getNumFmt().equals("chineseCountingThousand")){
            result = paragraph.getNumLevelText().replaceAll("%[1-9]+", ParaNumTool.numberToChineseCount(index));
        }else if(paragraph.getNumFmt().equals("chineseCounting")){
            result = paragraph.getNumLevelText().replaceAll("%[1-9]+",ParaNumTool.numberToChineseCount(index));
        } else if(paragraph.getNumFmt().equals("japaneseCounting")){
            result = paragraph.getNumLevelText().replaceAll("%[1-9]+",ParaNumTool.numberToChineseCount(index));
        }else if(paragraph.getNumFmt().equals("decimal")){
            result = paragraph.getNumLevelText().replaceAll("%[0-9]+",index+"");
        }else if(paragraph.getNumFmt().equals("decimalEnclosedCircle")){
            result = paragraph.getNumLevelText().replaceAll("%[0-9]+",index+"");
        }else if(paragraph.getNumFmt().equals("upperLetter")){
            result = paragraph.getNumLevelText().replaceAll("%[0-9]+",ParaNumTool.numToLetter(index).toUpperCase());
        }else if(paragraph.getNumFmt().equals("lowerLetter")){
            result = paragraph.getNumLevelText().replaceAll("%[0-9]+",ParaNumTool.numToLetter(index).toLowerCase());
        }else{
            //默认策略
            if(!paragraph.getNumLevelText().contains("%")){
                result = paragraph.getNumLevelText();
            }else{
                result = "";
            }
        }

        return result;

    }

    public static void main(String[] args) {
        String s = "%4)";
        String hh = s.replaceAll("%[1-9]+", "hh");
        System.out.println(hh);
    }
}
