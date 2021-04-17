package com.example.word.aspose;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.example.word.utils.TestFileUtil;

/**
 * @author liujunhui
 * @date 2021/2/1 9:13
 */
public class WordToPDF {

    public static void main(String[] args) throws Exception {
        String inPath = TestFileUtil.readDeskTopPath() + "丁权定制化练习手册-丁权-82316215.docx";
        String outPath = TestFileUtil.readDeskTopPath() + "1.pdf";
        convertWordToPDF(inPath, outPath);
    }

    /**
     * word转pdf
     * fixme 页眉线会和word不一样
     * @param inPath word文件路径
     * @param outPath pdf文件路径
     * @return
     */
    public static boolean convertWordToPDF(String inPath, String outPath) {
        Document document;
        try {
            document = new Document(inPath);
            document.save(outPath, SaveFormat.PDF);
            return true;
        } catch (Exception e) {
            return false;
        }
    }

}
