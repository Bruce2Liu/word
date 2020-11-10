package com.example.word.docx4j.template;

import com.example.word.constants.Constant;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.File;

/**
 * @author liujunhui
 * @date 2020/11/10 13:13
 */
public class Docx4jSimpleExample {
    public static void main(String[] args) throws Docx4JException {
        File docxFile = new File(Constant.outputPath + "docx4j.docx");

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        wordMLPackage.getMainDocumentPart().addParagraphOfText("hello world!");
        wordMLPackage.save(docxFile);
    }

}
