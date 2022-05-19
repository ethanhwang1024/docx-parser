package org.extract.text.tool;

import org.dom4j.Document;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;
import org.extract.pojo.ContentPojo;

import java.io.*;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @ClassName FileTool
 * @Description 存储文件
 * @Author WANGHAN756
 * @Date 2021/7/2 11:00
 * @Version 1.0
 **/
public class FileTool {
    public static String saveText(String outputFileDir,ContentPojo contentPojo,String absolutePath){
        if(outputFileDir==null){
            return "";
        }
        List<ContentPojo.Text> outList = contentPojo.getOutList();
        StringBuilder sb = new StringBuilder();
        outList.stream().forEach(x->{
            if(x.getElement_type().equals("text")||x.getElement_type().equals("title")){
                sb.append(x.getText().replaceAll("\n",""));
                sb.append("\n");
                sb.append("\n");
            }else if(x.getElement_type().equals("table")){
                //表格
                List<ContentPojo.Text.InnerCell> cells = x.getCells();
                Map<Integer, List<ContentPojo.Text.InnerCell>> collect = cells.stream().collect(Collectors.groupingBy(ContentPojo.Text.InnerCell::getRow_index));
                collect.entrySet().stream().forEach(y->{
                    List<ContentPojo.Text.InnerCell> rowCells = y.getValue();
                    for(ContentPojo.Text.InnerCell innerCell:rowCells){
                        sb.append(TextTool.unescape(innerCell.getText()).replaceAll("\n",""));
                        sb.append(",");
                    }
                    sb.append("\n");
                });
            }
        });
        FileWriter fwriter = null;
        try {
            String name = absolutePath.split("\\\\")[absolutePath.split("\\\\").length - 1].split("\\.")[0];
            // true表示不覆盖原来的内容，而是加到文件的后面。若要覆盖原来的内容，直接省略这个参数就好
            fwriter = new FileWriter(outputFileDir+"/"+name+".txt", false);
            fwriter.write(sb.toString());
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fwriter.flush();
                fwriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        return sb.toString();
    }

    public static String saveJson(String outputFileDir,ContentPojo contentPojo,String absolutePath){
        if(outputFileDir==null){
            return "";
        }
        String finalText = TextTool.toJson(contentPojo);
        FileWriter fwriter = null;
        try {
            String name = absolutePath.split("\\\\")[absolutePath.split("\\\\").length - 1].split("\\.")[0];
            // true表示不覆盖原来的内容，而是加到文件的后面。若要覆盖原来的内容，直接省略这个参数就好
            fwriter = new FileWriter(outputFileDir+"/"+name+".json", false);
            fwriter.write(finalText);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fwriter.flush();
                fwriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        return finalText;
    }

    public static void saveXML(String outputFileDir,Document dom,String absolutePath){
        if(outputFileDir==null||dom==null){
            return;
        }
        String name = absolutePath.split("\\\\")[absolutePath.split("\\\\").length - 1].split("\\.")[0];

        //设置生成xml格式
        OutputFormat format = OutputFormat.createPrettyPrint();
        // 设置编码格式
        format.setEncoding("UTF-8");
        File file = new File(outputFileDir+"/"+name+".xml");
        XMLWriter writer = null;
        try {
            writer = new XMLWriter(new FileOutputStream(file),format);
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        writer.setEscapeText(false); //关闭字符串中xml特殊字符转义
        try {
            writer.write(dom);
        } catch (IOException e) {
            e.printStackTrace();
        }
        try {
            writer.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static String saveHTML(String outputFileDir,ContentPojo contentPojo,String absolutePath){
        if(outputFileDir==null){
            return "";
        }
        List<ContentPojo.Text> outList = contentPojo.getOutList();
        StringBuilder sb = new StringBuilder();
        for(int i=0;i<outList.size();i++){
            ContentPojo.Text p = outList.get(i);
            String element_type = p.getElement_type();
            if("text".equals(element_type)||"title".equals(element_type)){
                //文本
                sb.append("<p>");
                sb.append(p.getText().replaceAll("\n","")).append("</p>");
                sb.append("<br/>");
            }else if(element_type.equals("table")){
                sb.append("<table border=\"1\">");
                List<ContentPojo.Text.InnerCell> cells = p.getCells();
                Map<Integer, List<ContentPojo.Text.InnerCell>> collect = cells.stream().collect(Collectors.groupingBy(ContentPojo.Text.InnerCell::getRow_index));
                collect.entrySet().stream().forEach(x->{
                    sb.append("<tr>").append("\n");
                    List<ContentPojo.Text.InnerCell> rowCells = x.getValue();
                    for(ContentPojo.Text.InnerCell innerCell:rowCells){
                        Integer row_span = innerCell.getRow_span();
                        Integer col_span = innerCell.getCol_span();
                        sb.append(String.format(" <td colspan=\"%s\" rowspan=\"%s\">", col_span , row_span));
                        sb.append(TextTool.unescape(innerCell.getText()));
                        sb.append("</td>");
                    }
                    sb.append("</tr>").append("\n");
                });
                sb.append("</table>").append("\n").append("<br/>");
            }else if(element_type.equals("pic")){
                //图片
                String text = p.getText();
                String format = String.format("<img src=\"%s\" />", text);
                sb.append(format);
                sb.append("<br/>");
            }
        }
        FileWriter fwriter = null;
        try {
            String name = absolutePath.split("\\\\")[absolutePath.split("\\\\").length - 1].split("\\.")[0];
            // true表示不覆盖原来的内容，而是加到文件的后面。若要覆盖原来的内容，直接省略这个参数就好
            fwriter = new FileWriter(outputFileDir+"/"+name+".html", false);
            fwriter.write(sb.toString());
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fwriter.flush();
                fwriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        return sb.toString();
    }
}
