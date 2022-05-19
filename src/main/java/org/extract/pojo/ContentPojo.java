package org.extract.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import org.extract.text.tool.TextTool;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;

import java.util.List;
import java.util.Map;

/**
 * @ClassName ContentPojo
 * @Description 输出json的映射pojo,注意：
 *              对于word文档来说，styleMap中对应docx解压后的style文件，存放
 *              部分样式，text中的List<StyleStruct>对应解压后的document文件中的
 *              样式信息，document中的样式要优先于style文件，而如果document中的文本
 *              没有指定styleId，将会默认使用style文件中的第一个样式。
 * @Author WANGHAN756
 * @Date 2021/4/21 9:20
 * @Version 1.0
 **/
@Data
@NoArgsConstructor
public class ContentPojo {
    public List<Text> outList;
    /**
     * WordStyleMap docx的style.xml里面的样式，当文本内容不带styleId时，会使用
     * styleMap中的第一个样式
     */
    public Map<String,WordStyleStruct> WordStyleMap;
    public ContentPojo(List<Text> outList){
        this.outList = outList;
    }
    //临时判断是否是ppt转成的pdf
    public Boolean isPptTransPDF;


    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class WordStyleStruct{
        private String text;
        private String styleId;
        private String styleName;
        private String pShd;
        //bA => beforeAutospacing,b => before, AA => afterAutospacing, A => after
        private String spacing;
        private String jc;
        private String fontAscii;
        private String fontEastAsia;
        private String fontHAnsi;
        private String fontCs;
        private String bold;
        private String bCs;
        private String i;
        private String iCs;
        private String smallCaps;
        private String strike;
        private String color;
        private String sz;
        private String szCs;
        private String t;
        private String u;
        private String shd;
    }



    @Data
    @NoArgsConstructor
    public static class Text{
        Integer page_number;
        String element_type; //title,text,table,pic
        String level;
        Boolean titleStart;
        String titlePrefix;
        String titleBody;

        String path;
        String styleId;
        //word文档有些段落开头的自动编号,这个字段标明
        String autoNumber;

        //段落加粗情况0->未加粗，1->部分加粗，2->全部加粗
        Integer bold;

        String text;
        Integer row_num;
        Integer col_num;
        Float xStart;
        Float yStart;
        Float width;
        Float height;

        Float pageHeight;
        Float pageWidth;


        //如果本元素有跨页，这里存储其合并的其他Text
        List<Text> crossPageList;


        //WordStyleStruct是word使用的样式
        List<WordStyleStruct> wordStyleStructs;


        List<InnerCell> cells;

        @Override
        public String toString() {
            return "Text{" +
                    "page_number=" + page_number +
                    ", element_type='" + element_type + '\'' +
                    ", level='" + level + '\'' +
                    ", text='" + text + '\'' +
                    ", row_num=" + row_num +
                    ", col_num=" + col_num +
                    ", cells=" + cells +
                    '}';
        }


        public Text(int page_number, String element_type, String text){
            this.page_number = page_number;
            this.element_type = element_type;
            this.text = text;
        }

        public Text(int page_number, String element_type, String level, String path,String text) {
            this.page_number = page_number;
            this.element_type = element_type;
            this.level = level;
            this.path = path;
            this.text = text;
        }

        public Text(int page_number,String element_type,int row_num,
                    int col_num,List<InnerCell> cells){
            this.page_number = page_number;
            this.element_type = element_type;
            this.row_num = row_num;
            this.col_num = col_num;
            this.cells = cells;
        }

        public Text(int page_number, String element_type,String level,String path,String text,Float xStart,Float yStart,
                    Float width,Float height){
            this.page_number = page_number;
            this.element_type = element_type;
            this.level = level;
            this.path = path;
            this.text = text;
            this.xStart = xStart;
            this.yStart = yStart;
            this.width = width;
            this.height = height;
        }

        public Text(int page_number, String element_type, String text,Float xStart,Float yStart,
                    Float width,Float height){
            this.page_number = page_number;
            this.element_type = element_type;
            this.text = text;
            this.xStart = xStart;
            this.yStart = yStart;
            this.width = width;
            this.height = height;
        }

        public Text(int page_number,String element_type,int row_num,
                    int col_num,List<InnerCell> cells,Float xStart,Float yStart,
                    Float width,Float height){
            this.page_number = page_number;
            this.element_type = element_type;
            this.row_num = row_num;
            this.col_num = col_num;
            this.cells = cells;
            this.xStart = xStart;
            this.yStart = yStart;
            this.width = width;
            this.height = height;
        }
        @Data
        public static class InnerCell{
            String text;
            Integer row_index;
            Integer col_index;
            Integer row_span;
            Integer col_span;

            //word表格解析临时使用的两个变量，序列化时需要排除
            Integer trueColIndex;
            CTVMerge vMerge;
            //目前支持跨两页的表格合并，因此对pdf有必要有一个字段表示这个单元格在下一页
            Boolean isNextPage;
            Float xStart;
            Float yStart;
            Float width;
            Float height;


            public InnerCell(String text,Integer row_index,Integer col_index,
                             Integer row_span,Integer col_span){
                this.text = text;
                this.row_index = row_index;
                this.col_index = col_index;
                this.row_span = row_span;
                this.col_span = col_span;
            }

            public InnerCell() {

            }


            @Override
            public String toString() {
                return "InnerCell{" +
                        "text='" + text + '\'' +
                        ", row_index=" + row_index +
                        ", col_index=" + col_index +
                        ", row_span=" + row_span +
                        ", col_span=" + col_span +
                        '}';
            }
        }
    }
}
