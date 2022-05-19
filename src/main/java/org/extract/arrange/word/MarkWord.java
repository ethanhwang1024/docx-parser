package org.extract.arrange.word;

import org.apache.commons.collections4.CollectionUtils;
import org.extract.pojo.ContentPojo;
import org.extract.pojo.MarkPojo;
import org.extract.text.tool.SettingReader;
import org.extract.text.tool.TextTool;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @ClassName MarkWord
 * @Description 根据配置文件标记Word段落中的标题
 * @Author WANGHAN756
 * @Date 2021/7/2 15:23
 * @Version 1.0
 **/
public class MarkWord {
    /**
     * 根据配置文件对标题进行标记
     * @param contentPojo contentPojo
     * @param markPojo markPojo 对应配置文件中的标记json
     */
    public static void markTitle(ContentPojo contentPojo,MarkPojo markPojo){
        if(markPojo==null){
            markPojo = SettingReader.getDocMark();
        }
        List<MarkPojo.TitlePattern> titlePatterns = markPojo.getTitlePatterns();
        //按order进行排序
        List<MarkPojo.TitlePattern> sortedTitlePatterns = titlePatterns.stream().filter(x->x.getOrder()!=null).sorted(new Comparator<MarkPojo.TitlePattern>() {
            @Override
            public int compare(MarkPojo.TitlePattern o1, MarkPojo.TitlePattern o2) {
                return o1.getOrder() - o2.getOrder();
            }
        }).collect(Collectors.toList());

        List<ContentPojo.Text> outList = contentPojo.getOutList();

        for(int i=0;i<outList.size();i++){
            ContentPojo.Text p = outList.get(i);
            if(p.getElement_type().equals("table")||p.getElement_type().equals("pic")){
                continue;
            }
            List<ContentPojo.WordStyleStruct> styles = p.getWordStyleStructs();

            if(styles!=null){
                for(int j=0;j<sortedTitlePatterns.size();j++){
                    MarkPojo.TitlePattern titlePattern = sortedTitlePatterns.get(j);
                    String bold = titlePattern.getBold();
                    List<Integer> boldStatuses = new ArrayList<>();
                    String pattern = titlePattern.getPattern();
                    String firstPattern = titlePattern.getFirstPattern();
                    Float level = titlePattern.getLevel();
                    boldStatuses.add(1);boldStatuses.add(2);boldStatuses.add(0);
                    int boldStatus = 0;
                    if(bold!=null){
                        boldStatuses = Arrays.stream(bold.split(",")).map(Integer::parseInt).collect(Collectors.toList());
                        boldStatus = verifyBold(styles);
                    }
                    if(boldStatuses.contains(boldStatus)){
                        //先看整体是否是符合要求的
                        if(p.getText().matches(pattern)){
                            p.setElement_type("title");
                            p.setLevel(level+"");
                            //捕获组抽取标题部分
                            Pattern pa = Pattern.compile(pattern);
                            Matcher m = pa.matcher(p.getText());
                            if(m.find()){
                                if(m.groupCount()==2){
                                    p.setTitlePrefix(m.group(1));
                                    p.setTitleBody(m.group(2));

                                }
                            }

                            if(firstPattern!=null){
                                if(p.getText().matches(firstPattern)){
                                    //如果是本级标题的初始标题
                                    p.setTitleStart(true);
                                }else{
                                    p.setTitleStart(false);
                                }
                            }
                        }
                    }
                }
            }
        }
        //过滤掉目录部分，找到一级标题的最初的位置，如果有多个，就定位到第最后一个的位置，那么前面如果有任何标题就设置为文本
        Optional<MarkPojo.TitlePattern> first = titlePatterns.stream().filter(x -> x.getLevel() == 1f).findFirst();
        if(first.isPresent()){
            List<Integer> firstHeaderList = new ArrayList<>();
            for(int i=0;i<outList.size();i++){
                ContentPojo.Text p = outList.get(i);
                String element_type = p.getElement_type();
                if(element_type.equals("title")){
                    String level = p.getLevel();
                    Boolean titleStart = p.getTitleStart();
                    if(titleStart!=null&&titleStart&&level.equals("1.0")){
                        //第一级别标题
                        firstHeaderList.add(i);
                    }
                }
            }
            if(!CollectionUtils.isEmpty(firstHeaderList)){
                Integer firstHeaderIndex = firstHeaderList.get(firstHeaderList.size() - 1);
                for(int i=0;i<firstHeaderIndex;i++){
                    ContentPojo.Text p = outList.get(i);
                    if(p.getElement_type().equals("title")){
                        p.setElement_type("text");
                        p.setLevel(null);
                        p.setTitleStart(null);
                    }
                }
            }
        }

    }
    /**
     *  多种情况是加粗的
     *  1.fontName:带SimHei，如ABCDEE+SimHei
     *  2.fontWeight大于400
     *  3.renderingMode为FILL_STROKE
     *  注意只考虑中英文文本的加粗情况,特殊符号跳过
     *  0->不包含，1->包含部分，2->全部加粗
     */
    private static Integer verifyBold(List<ContentPojo.WordStyleStruct> styles){
        int countAll = 0;
        int countBold = 0;
        for(int i=0;i<styles.size();i++){
            ContentPojo.WordStyleStruct style = styles.get(i);
            if(style!=null){
                String text = style.getText();
                if(TextTool.isContainEnglish(text)||TextTool.isContainChinese(text)){
                    countAll++;
                    String bold = style.getBold();
                    if(bold!=null){
                        countBold++;
                    }
                }
            }
        }
        if(countAll!=0){
            if(countBold==countAll){
                return 2;
            }else if(countBold>=1){
                return 1;
            }else{
                return 0;
            }
        }else{
            return 0;
        }
    }

}
