import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.extract.arrange.word.MarkWord;
import org.extract.arrange.word.WordXml;
import org.extract.core.WordParserEngine;
import org.extract.pojo.ContentPojo;
import org.extract.pojo.Tu;
import org.extract.text.tool.FileTool;
import org.junit.Test;

import java.io.FileInputStream;
import java.util.List;

@SuppressWarnings("all")
public class SingleTest {
    //
    private static String inputFilePath = SingleTest.class.getClassLoader().getResource("sample.docx").getPath();;
    //The path needs to exist in advance
    private static String outputFilePathDir = "E:\\年报\\output";
    //The path needs to exist in advance
    private static String picSavePath = "E:\\pic";

    @Test
    public void parsing() throws Exception {
        XWPFDocument pf = new XWPFDocument(new FileInputStream(inputFilePath));
        parsing(pf,"sample",outputFilePathDir);
    }

    public static String parsing(XWPFDocument document,String fileName,String outputFileDir){

        ContentPojo contentPojo = WordParserEngine.parsing(document,picSavePath,null);
        String json = FileTool.saveJson(outputFileDir, contentPojo, fileName);
        FileTool.saveHTML(outputFileDir,contentPojo,fileName);
        FileTool.saveText(outputFileDir,contentPojo,fileName);
        return json;
    }
}
