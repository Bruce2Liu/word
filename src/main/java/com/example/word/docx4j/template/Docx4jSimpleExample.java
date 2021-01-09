package com.example.word.docx4j.template;

import com.example.word.constants.Constant;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author liujunhui
 * @date 2020/11/10 13:13
 */
public class Docx4jSimpleExample {

    public static void main(String[] args) throws Docx4JException {
        File docxFile = new File(Constant.outputPath + "docx4j.docx");
        try {
            Files.createDirectories(Paths.get(docxFile.getParent()));
        } catch (IOException e) {
            e.printStackTrace();
        }

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        wordMLPackage.getMainDocumentPart().addParagraphOfText("hello world!");
        addStyledPText(wordMLPackage);
        wordMLPackage.save(docxFile);
    }

    /**
     * 加载读入word文件
     */
    public static void loadWord(String docxFilePath) {
        try {
            WordprocessingMLPackage template = WordprocessingMLPackage.load(new File("c:\\a.docx"));
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }

    /**
     * 文本添加样式
     */
    public static void addStyledPText(WordprocessingMLPackage wordMLPackage) {
        wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("Title", "标题样式");
        wordMLPackage.getMainDocumentPart().addStyledParagraphOfText("SubTitle",  "SubTitle样式");
    }

    public static void addTable() {

    }

}
