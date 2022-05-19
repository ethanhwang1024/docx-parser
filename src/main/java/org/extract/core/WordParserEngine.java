package org.extract.core;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xwpf.usermodel.*;
import org.extract.arrange.word.MarkWord;
import org.extract.arrange.word.WordXml;
import org.extract.pojo.ContentPojo;
import org.extract.pojo.ExtractPojo;
import org.extract.pojo.MarkPojo;
import org.extract.pojo.Tu;
import org.extract.text.analyser.word.ParaNumHandler;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.*;

/**
 * @ClassName WordParserEngine
 * @Description word文档转换为ContentPojo
 * @Author WANGHAN756
 * @Date 2021/5/10 17:37
 * @Version 1.0
 **/
public class WordParserEngine {
    private static Field listStyleField;
    
    static {
        try {
            Class<?> XWPFStylesC = Class.forName("org.apache.poi.xwpf.usermodel.XWPFStyles");
            listStyleField = XWPFStylesC.getDeclaredField("listStyle");
            listStyleField.setAccessible(true);
        } catch (ClassNotFoundException | NoSuchFieldException e) {
            e.printStackTrace();
        }
    }



    public static ContentPojo parsing(XWPFDocument document,String attachSavePath,String urlPath){

        XWPFStyles xwpfStyles = document.getStyles();

        List<XWPFStyle> styles = new ArrayList<>();
        try {
            styles = (List<XWPFStyle>)listStyleField.get(xwpfStyles);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        Map<String,ContentPojo.WordStyleStruct> styleMap = new LinkedHashMap<>();
        for(XWPFStyle x:styles){
            ContentPojo.WordStyleStruct styleStruct = new ContentPojo.WordStyleStruct();
            String styleId = x.getStyleId();
            styleStruct.setStyleId(styleId);
            String name = x.getName();
            styleStruct.setStyleName(name);
            CTStyle ctStyle = x.getCTStyle();
            if(ctStyle!=null){
                CTPPr pPr = ctStyle.getPPr();
                if(pPr!=null){
                    CTShd pShd = pPr.getShd();
                    if(pShd!=null){
                        STShd.Enum val = pShd.getVal();
                        if(val!=null){
                            styleStruct.setPShd(val.toString());
                        }else{
                            styleStruct.setPShd("");
                        }
                    }
                    CTSpacing spacing = pPr.getSpacing();
                    if(spacing!=null){
                        Object beforeAutospacing = spacing.getBeforeAutospacing();
                        Object before = spacing.getBefore();
                        Object after = spacing.getAfter();
                        Object afterAutospacing = spacing.getAfterAutospacing();
                        String pos = "bA:"+beforeAutospacing+",b:"+before+",AA:"+afterAutospacing+
                                ",A:"+after;
                        styleStruct.setSpacing(pos);
                    }
                    CTJc jc = pPr.getJc();
                    if(jc!=null){
                        STJc.Enum val = jc.getVal();
                        if(val!=null){
                            styleStruct.setJc(val.toString());
                        }else{
                            styleStruct.setJc("");
                        }
                    }
                }

                CTRPr rPr = ctStyle.getRPr();
                if(rPr!=null){
                    CTFonts rFonts = rPr.getRFonts();
                    //一般来说size为1
                    if(rFonts!=null){
                        String ascii = rFonts.getAscii();
                        styleStruct.setFontAscii(ascii);
                        String cs = rFonts.getCs();
                        styleStruct.setFontCs(cs);
                        String eastAsia = rFonts.getEastAsia();
                        styleStruct.setFontEastAsia(eastAsia);
                        String hAnsi = rFonts.getHAnsi();
                        styleStruct.setFontHAnsi(hAnsi);
                    }
                    //粗体情况，如果val为0，其实也不是粗体
                    CTOnOff b = rPr.getB();
                    if(b!=null){
                        Object val = b.getVal();
                        if(val!=null){
                            styleStruct.setBold(val.toString());
                        }else{
                            styleStruct.setBold("");
                        }
                    }

                    CTOnOff bCs = rPr.getBCs();
                    if(bCs!=null){
                        Object val = bCs.getVal();
                        if(val!=null){
                            styleStruct.setBCs(val.toString());
                        }else{
                            styleStruct.setBCs("");
                        }
                    }

                    CTOnOff ctOnOff = rPr.getI();
                    if(ctOnOff!=null){
                        Object val = ctOnOff.getVal();
                        if(val!=null){
                            styleStruct.setI(val.toString());
                        }else{
                            styleStruct.setI("");
                        }
                    }


                    CTOnOff iCs = rPr.getICs();
                    if(iCs!=null){
                        Object val = iCs.getVal();
                        if(val!=null){
                            styleStruct.setICs(val.toString());
                        }else{
                            styleStruct.setICs("");
                        }
                    }

                    CTOnOff smallCaps = rPr.getSmallCaps();
                    if(smallCaps!=null){
                        Object val = smallCaps.getVal();
                        if(val!=null){
                            styleStruct.setSmallCaps(val.toString());
                        }else{
                            styleStruct.setSmallCaps("");
                        }
                    }

                    CTOnOff strike = rPr.getStrike();
                    if(strike!=null){
                        Object val = strike.getVal();
                        if(val!=null){
                            styleStruct.setStrike(val.toString());
                        }else{
                            styleStruct.setStrike("");
                        }
                    }

                    CTColor color = rPr.getColor();
                    if(color!=null){
                        STThemeColor.Enum themeColor = color.getThemeColor();
                        if(themeColor!=null){
                            styleStruct.setColor(themeColor.toString());
                        }else{
                            styleStruct.setColor("");
                        }
                    }


                    CTHpsMeasure sz = rPr.getSz();
                    if(sz!=null){
                        Object val = sz.getVal();
                        if(val!=null){
                            styleStruct.setSz(val.toString());
                        }else{
                            styleStruct.setSz("");
                        }
                    }

                    CTHpsMeasure szCs = rPr.getSzCs();
                    if(szCs!=null){
                        Object val = szCs.getVal();
                        if(val!=null){
                            styleStruct.setSzCs(val.toString());
                        }else{
                            styleStruct.setSzCs("");
                        }
                    }

                    CTUnderline u = rPr.getU();
                    if(u!=null){
                        STUnderline.Enum anEnum =u.getVal();
                        if(anEnum!=null){
                            styleStruct.setU(anEnum.toString());
                        }else{
                            styleStruct.setU("");
                        }
                    }
                    CTShd shd = rPr.getShd();
                    if(shd!=null){
                        STShd.Enum stshdEnum = shd.getVal();
                        if(stshdEnum!=null){
                            styleStruct.setShd(stshdEnum.toString());
                        }else{
                            styleStruct.setShd("");
                        }
                    }


                }
            }
            styleMap.put(styleId,styleStruct);
        }

        ContentPojo contentPojo = new ContentPojo();
        contentPojo.setWordStyleMap(styleMap);

        ParaNumHandler paraNumHandler = new ParaNumHandler();
        List<ContentPojo.Text> outList = new ArrayList<>();
        try {
            // 获取word中的所有段落与表格
            List<IBodyElement> elements = document.getBodyElements();
            int pageNumber = 1;
            StringBuilder sb = new StringBuilder();
            for (IBodyElement element : elements) {
                // 段落
                if (element instanceof XWPFParagraph) {
                    XWPFParagraph para = (XWPFParagraph) element;

                    //获得自动段落开始自动编码的数字或者中文,类似(1),(一)
                    String autoNumber = paraNumHandler.getParagraphNum(para);
                    if(autoNumber!=null){
                        sb.append(autoNumber);
                    }


                    CTP ctp = para.getCTP();
                    List<CTR> rList = ctp.getRList();
                    List<ContentPojo.WordStyleStruct> styleStructs = new ArrayList<>();
                    for(int i=0;i<rList.size();i++){
                        ContentPojo.WordStyleStruct styleStruct = new ContentPojo.WordStyleStruct();
                        CTR ctr = rList.get(i);
                        CTRPr rPr = ctr.getRPr();
                        List<CTText> tList = ctr.getTList();
                        if(tList.size()!=0){
                            styleStruct.setText(tList.get(0).getStringValue());
                        }
                        if(rPr!=null){
                            CTFonts rFonts = rPr.getRFonts();
                            //一般来说size为1
                            if(rFonts!=null){
                                String ascii = rFonts.getAscii();
                                styleStruct.setFontAscii(ascii);
                                String cs = rFonts.getCs();
                                styleStruct.setFontCs(cs);
                                String eastAsia = rFonts.getEastAsia();
                                styleStruct.setFontEastAsia(eastAsia);
                                String hAnsi = rFonts.getHAnsi();
                                styleStruct.setFontHAnsi(hAnsi);
                            }
                            //粗体情况，如果val为0，其实也不是粗体
                            CTOnOff b = rPr.getB();
                            if(b!=null){
                                Object val = b.getVal();
                                if(val!=null){
                                    styleStruct.setBold(val.toString());
                                }else{
                                    styleStruct.setBold("");
                                }
                            }

                            CTOnOff bCs = rPr.getBCs();
                            if(bCs!=null){
                                Object val = bCs.getVal();
                                if(val!=null){
                                    styleStruct.setBCs(val.toString());
                                }else{
                                    styleStruct.setBCs("");
                                }
                            }

                            CTOnOff ctOnOff = rPr.getI();
                            if(ctOnOff!=null){
                                Object val = ctOnOff.getVal();
                                if(val!=null){
                                    styleStruct.setI(val.toString());
                                }else{
                                    styleStruct.setI("");
                                }
                            }


                            CTOnOff iCs = rPr.getICs();
                            if(iCs!=null){
                                Object val = iCs.getVal();
                                if(val!=null){
                                    styleStruct.setICs(val.toString());
                                }else{
                                    styleStruct.setICs("");
                                }
                            }

                            CTOnOff smallCaps = rPr.getSmallCaps();
                            if(smallCaps!=null){
                                Object val = smallCaps.getVal();
                                if(val!=null){
                                    styleStruct.setSmallCaps(val.toString());
                                }else{
                                    styleStruct.setSmallCaps("");
                                }
                            }

                            CTOnOff strike = rPr.getStrike();
                            if(strike!=null){
                                Object val = strike.getVal();
                                if(val!=null){
                                    styleStruct.setStrike(val.toString());
                                }else{
                                    styleStruct.setStrike("");
                                }
                            }

                            CTColor color = rPr.getColor();
                            if(color!=null){
                                STThemeColor.Enum themeColor = color.getThemeColor();
                                if(themeColor!=null){
                                    styleStruct.setColor(themeColor.toString());
                                }else{
                                    styleStruct.setColor("");
                                }
                            }


                            CTHpsMeasure sz = rPr.getSz();
                            if(sz!=null){
                                Object val = sz.getVal();
                                if(val!=null){
                                    styleStruct.setSz(val.toString());
                                }else{
                                    styleStruct.setSz("");
                                }
                            }

                            CTHpsMeasure szCs = rPr.getSzCs();
                            if(szCs!=null){
                                Object val = szCs.getVal();
                                if(val!=null){
                                    styleStruct.setSzCs(val.toString());
                                }else{
                                    styleStruct.setSzCs("");
                                }
                            }

                            CTUnderline u = rPr.getU();
                            if(u!=null){
                                STUnderline.Enum anEnum =u.getVal();
                                if(anEnum!=null){
                                    styleStruct.setU(anEnum.toString());
                                }else{
                                    styleStruct.setU("");
                                }
                            }
                            CTShd shd = rPr.getShd();
                            if(shd!=null){
                                STShd.Enum stshdEnum = shd.getVal();
                                if(stshdEnum!=null){
                                    styleStruct.setShd(stshdEnum.toString());
                                }else{
                                    styleStruct.setShd("");
                                }
                            }

                        }
                        CTPPr pPr = ctp.getPPr();
                        if(pPr!=null){
                            CTShd pShd = pPr.getShd();
                            if(pShd!=null){
                                STShd.Enum val = pShd.getVal();
                                if(val!=null){
                                    styleStruct.setPShd(val.toString());
                                }else{
                                    styleStruct.setPShd("");
                                }
                            }
                            CTSpacing spacing = pPr.getSpacing();
                            if(spacing!=null){
                                Object beforeAutospacing = spacing.getBeforeAutospacing();
                                Object before = spacing.getBefore();
                                Object after = spacing.getAfter();
                                Object afterAutospacing = spacing.getAfterAutospacing();
                                String pos = "bA:"+beforeAutospacing+",b:"+before+",AA:"+afterAutospacing+
                                        ",A:"+after;
                                styleStruct.setSpacing(pos);
                            }
                            CTJc jc = pPr.getJc();
                            if(jc!=null){
                                STJc.Enum val = jc.getVal();
                                if(val!=null){
                                    styleStruct.setJc(val.toString());
                                }else{
                                    styleStruct.setJc("");
                                }
                            }
                        }
                        styleStructs.add(styleStruct);
                    }

                    String styleID = para.getStyleID();
                    String eleType = "text";
                    List<XWPFRun> runs = para.getRuns();
                    for (XWPFRun run : runs) {
                        Node node = run.getCTR().getDomNode();
                        String picType = "drawing";
                        Node picNode = getChildNode(node, "w:drawing");
                        if(picNode==null){
                            picNode = getChildNode(node, "w:pict");
                            picType = "pict";
                        }
                        if (picNode == null) {
                            sb.append(run.text());
                        }else{
                            //发现有图片，断开
                            ContentPojo.Text p = new ContentPojo.Text();
                            p.setWordStyleStructs(styleStructs);
                            p.setPage_number(pageNumber);
                            p.setElement_type(eleType);
                            p.setAutoNumber(autoNumber);
                            p.setText(sb.toString());
                            p.setStyleId(styleID);
                            outList.add(p);
                            sb.delete(0,sb.length());
                            String pathUrl = processPic(picType,picNode, document, attachSavePath, urlPath);
                            //存储图片地址
                            if(!StringUtils.isEmpty(pathUrl)){
                                ContentPojo.Text picP = new ContentPojo.Text();
                                picP.setElement_type("pic");
                                picP.setText(pathUrl);
                                outList.add(picP);
                            }
                        }
                    }
                    if(sb.length()>=1){
                        ContentPojo.Text p = new ContentPojo.Text();
                        p.setWordStyleStructs(styleStructs);
                        p.setPage_number(pageNumber);
                        p.setElement_type(eleType);

                        p.setAutoNumber(autoNumber);
                        p.setText(sb.toString());

                        p.setStyleId(styleID);
                        outList.add(p);
                        sb.delete(0,sb.length());
                    }

                    if(para.getCTP().toString().contains("lastRenderedPageBreak")){
                        pageNumber++;
                    }
                }
                // 表格
                else if (element instanceof XWPFTable) {
                    int savedPageNumber = pageNumber;
                    //如果还有段落残留，先把段落处理完毕
                    if(sb.length()>1){
                        outList.add(new ContentPojo.Text(pageNumber,"text",sb.toString()));
                        sb.delete(0,sb.length());
                    }

                    XWPFTable table = (XWPFTable) element;

                    List<XWPFTableRow> rows = table.getRows();

                    //外层List是不同行，内层list是每行的单元格
                    List<List<ContentPojo.Text.InnerCell>> innerCells = new ArrayList<>();

//                    List<ContentPojo.Text.InnerCell> innerCells = new ArrayList<>();

                    int rowNum = rows.size();
                    int colNum = 0;

                    for(int i=0;i<rows.size();i++){
                        //每一行都向innerCells先放入一个空的列表
                        innerCells.add(new ArrayList<>());

                        XWPFTableRow row = rows.get(i);

                        List<Tu.Tuple3<Integer,String,XWPFTableCell>> pairCell = new ArrayList<>();


//                        CTTc[] tcArray = row.getCtRow().getTcArray();
                        boolean isPreCellHasRightBorder = false;


                        for (CTTc tableCell : row.getCtRow().getTcArray()) {
                            XWPFTableCell cell = new XWPFTableCell(tableCell, row, table.getBody());
                            String cellText = getCellText(cell);
                            try{
                                CTBorder left = tableCell.getTcPr().getTcBorders().getLeft();
                                STBorder.Enum valr = null;
                                try{
                                    CTBorder right = tableCell.getTcPr().getTcBorders().getRight();
                                    valr = right.getVal();
                                }catch (Exception e){
                                    //todo 用日志进行打印
//                                    e.printStackTrace();
                                }

                                STBorder.Enum val = left.getVal();
//                            CTTblWidth tcW = tableCell.getTcPr().getTcW();
                                //如果本单元格左侧有边框，或者左侧的一个单元格有右边框
                                if(val!=STBorder.NIL||isPreCellHasRightBorder){
                                    pairCell.add(new Tu.Tuple3<>(0,cellText,new XWPFTableCell(tableCell, row, table.getBody())));
                                }else{
                                    //如果本单元格左侧就是不存在边框
                                    if(pairCell.size()!=0){
                                        pairCell.get(pairCell.size()-1).setValue1(pairCell.get(pairCell.size()-1).getValue1()+1);
                                        pairCell.get(pairCell.size()-1).setValue2(pairCell.get(pairCell.size()-1).getValue2()+" "+cellText);
                                    }
                                }
                                if(valr!=null&&valr!=STBorder.NIL){
                                    isPreCellHasRightBorder = true;
                                }else{
                                    isPreCellHasRightBorder = false;
                                }
                            }catch (Exception e){
                                pairCell.add(new Tu.Tuple3<>(0,cellText,new XWPFTableCell(tableCell, row, table.getBody())));
                            }

                        }


                        colNum = Math.max(colNum,pairCell.size());

                        int trueCol =0;
                        for(int j=0;j<pairCell.size();j++){
                            Integer addColSpan = pairCell.get(j).getValue1();
                            String cellText = pairCell.get(j).getValue2();
                            XWPFTableCell cell = pairCell.get(j).getValue3();
                            CTTc ctTc = cell.getCTTc();
                            if(ctTc.toString().contains("lastRenderedPageBreak")){
                                pageNumber++;
                            }
                            CTDecimalNumber gridSpan = cell.getCTTc().getTcPr().getGridSpan();
                            CTVMerge vMerge = cell.getCTTc().getTcPr().getVMerge();
                            ContentPojo.Text.InnerCell innerCell = new ContentPojo.Text.InnerCell();

                            innerCell.setText(cellText);
                            innerCell.setRow_index(i+1);
                            innerCell.setCol_index(j+1);
                            innerCell.setVMerge(vMerge);

                            if(gridSpan==null){
                                innerCell.setCol_span(1+addColSpan);
                                innerCell.setTrueColIndex(trueCol+1);
                                trueCol += (1+addColSpan);
                            }else{
                                BigInteger val = gridSpan.getVal();
                                innerCell.setCol_span(val.intValue()+addColSpan);
                                innerCell.setTrueColIndex(trueCol+1);
                                trueCol += (val.intValue()+addColSpan);
                            }

                            innerCells.get(innerCell.getRow_index()-1).add(innerCell);
                        }
                    }

                    List<ContentPojo.Text.InnerCell> combineList = new ArrayList<>();

                    List<ContentPojo.Text.InnerCell> waitRemoveCell = new ArrayList<>();
                    for(int i=0;i<innerCells.size()-1;i++){
                        List<ContentPojo.Text.InnerCell> cells = innerCells.get(i);
                        for(int j=0;j<cells.size();j++){
                            ContentPojo.Text.InnerCell innerCell = cells.get(j);
                            CTVMerge vMerge = innerCell.getVMerge();
                            if(vMerge==null) {
                                innerCell.setRow_span(1);
                            }else{
                                STMerge.Enum val = vMerge.getVal();
                                if(val==null){
                                    innerCell.setRow_span(1);
                                    continue;
                                }else{
                                    //这个单元格是向下合并的初始单元格
                                    int count = 1;
                                    Integer trueColIndex = innerCell.getTrueColIndex();
                                    a:for(int k=i+1;k<innerCells.size();k++){
                                        List<ContentPojo.Text.InnerCell> tmpCells = innerCells.get(k);
                                        for(ContentPojo.Text.InnerCell tmpCell:tmpCells){
                                            Integer tmpCellTrueColIndex = tmpCell.getTrueColIndex();
                                            if(tmpCellTrueColIndex.equals(trueColIndex)){
                                                CTVMerge tmpVMerge = tmpCell.getVMerge();
                                                if(tmpVMerge!=null){
                                                    if(tmpVMerge.getVal()!=null){
                                                        break a;
                                                    }
                                                    innerCell.setText(innerCell.getText()+tmpCell.getText());
                                                    //被合并的单元格消除
                                                    waitRemoveCell.add(tmpCell);
                                                    count++;
                                                }else{
                                                    break;
                                                }
                                            }
                                        }
                                        innerCell.setRow_span(count);
                                    }
                                }
                            }
                        }
                        combineList.addAll(cells);
                    }
                    if(innerCells.size()>=1){
                        List<ContentPojo.Text.InnerCell> innerCellList = innerCells.get(innerCells.size() - 1);
                        for(ContentPojo.Text.InnerCell innerCell:innerCellList){
                            innerCell.setRow_span(1);
                        }
                        combineList.addAll(innerCells.get(innerCells.size()-1));
                    }

                    for(ContentPojo.Text.InnerCell tmpCell:waitRemoveCell){
                        combineList.remove(tmpCell);
                    }
                    denseCols(combineList);
                    outList.add(new ContentPojo.Text(savedPageNumber, "table", rowNum, colNum,combineList));
                }


            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                document.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        contentPojo.setOutList(outList);
        return contentPojo;
    }

    /**
     * 直接获取标题和内容的方法
     * @param document 文档
     * @param extractPojo extractPojo对应json文件中的抽取json，如果这里传null，那么会启用
     *                    默认的配置文件，但是当本jar包嵌入到其他项目里的时候，建议在那个项目
     *                    里面放配置文件，然后自己读取转换成extractPojo传入
     * @param markPojo markPojo对应json文件中的标记json，同extractPojo
     * @return 标题和对应内容
     * @throws IOException
     */
    public static List<Tu.Tuple2<String, String>> getExtractContent(XWPFDocument document,String attachSavePath,String urlPath,ExtractPojo extractPojo, MarkPojo markPojo){
        ContentPojo contentPojo = parsing(document,attachSavePath,urlPath);
        MarkWord.markTitle(contentPojo,markPojo);
        org.dom4j.Document doc1 = WordXml.buildXml(contentPojo);
        List<Tu.Tuple2<String, String>> extract = WordXml.extract(doc1, contentPojo,extractPojo);
        return extract;
    }

    /**
     * 处理图片
     * @param picType drawing或pict
     * @param node 节点
     * @param document 文档
     * @param attachSavePath 需要保存的路径，如果传null，将被保存在tmp路径下
     * @param urlPath url路径，会返回这个路径的拼接路径，如果传null，将会返回绝对路径
     * @return url路径
     * @throws IOException
     */
    private static String processPic(String picType,Node node,XWPFDocument document,String attachSavePath,String urlPath) throws IOException {
        String id = "";
        if(picType.equals("drawing")){
            Node blipNode = getChildNode(node, "a:blip");
            if(blipNode==null){
                return null;
            }
            NamedNodeMap blipAttrs = blipNode.getAttributes();
            id = blipAttrs.getNamedItem("r:embed").getNodeValue();
        }else if(picType.equals("pict")){
            Node dataNode = getChildNode(node, "v:imagedata");
            if(dataNode==null){
                return null;
            }
            NamedNodeMap shapeAttrs = dataNode.getAttributes();
            id = shapeAttrs.getNamedItem("r:id").getNodeValue();
        }
        if(id.equals("")){
            return null;
        }
        // 获取图片信息
        PackagePart part = document.getPartById(id);

        String suffix = getSuffix(part.getPartName().getName());
        String picName= UUID.randomUUID().toString().replace("-", "")  + suffix;
        File picFile = null;
        //创建临时文件
        picFile = createTmpFileWithName(picName,attachSavePath);
        String curUrlPath = "";
        if(urlPath==null){
            curUrlPath = picFile.getAbsolutePath();
        }else{
            curUrlPath = urlPath + "/"+ picName;
        }
        InputStream is = part.getInputStream();
        Files.copy(is, picFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        return curUrlPath;
    }

    private static String getSuffix(String fileName) {
        if(fileName==null){
            return "";
        }
        if (fileName.contains(".")) {
            String suffix = fileName.substring(fileName.lastIndexOf("."));
            return suffix.toLowerCase();
        }
        return "";
    }

    private static File getTmpDir() {
        String projectPath = System.getProperty("user.dir") + File.separator + "temp";
        File file = new File(projectPath);
        if (!file.exists()) {
            file.mkdirs();
        }
        return file;
    }

    private static File createTmpFileWithName(String fileName,String attachSavePath) throws IOException {
        if(attachSavePath==null){
            File file = new File(getTmpDir(), fileName);
            if (!file.exists()) {
                file.createNewFile();
            }
            return file;
        }else{
            File file = new File(attachSavePath, fileName);
            if (!file.exists()) {
                file.createNewFile();
            }
            return file;
        }

    }


    private static Node getChildNode(Node node, String nodeName) {
        if (!node.hasChildNodes()) {
            return null;
        }
        NodeList childNodes = node.getChildNodes();
        for (int i = 0; i < childNodes.getLength(); i++) {
            Node childNode = childNodes.item(i);
            if (nodeName.equals(childNode.getNodeName())) {
                return childNode;
            }
            childNode = getChildNode(childNode, nodeName);
            if (childNode != null) {
                return childNode;
            }
        }
        return null;
    }


    private static String getCellText(XWPFTableCell cell){
        StringBuilder sb = new StringBuilder();
        List<XWPFParagraph> paragraphs = cell.getParagraphs();

        for(XWPFParagraph para:paragraphs){
            List<XWPFRun> runs = para.getRuns();
            if (runs.size() == 0) {
                //回车换行
                sb.append("\n");
            }else{
                for (XWPFRun run : runs) {
                    sb.append(run.text());
                }
            }
        }
        return sb.toString();
    }


    private static void denseCols(List<ContentPojo.Text.InnerCell> cells){
        int curRow = 0;
        int preCol = 0;
        for (ContentPojo.Text.InnerCell cell : cells) {
            Integer rowIndex = cell.getRow_index();
            if (curRow == 0) {
                curRow = rowIndex;
                preCol = 1;
                cell.setCol_index(preCol);
            } else {
                if (rowIndex == curRow) {
                    //说明是某一行的延续
                    cell.setCol_index(++preCol);
                } else if (rowIndex > curRow) {
                    //说明某一行新起
                    curRow = rowIndex;
                    preCol = 1;
                    cell.setCol_index(preCol);
                }
            }
        }
    }
}
