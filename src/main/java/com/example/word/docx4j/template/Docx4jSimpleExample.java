package com.example.word.docx4j.template;

import com.example.word.constants.Constant;
import com.example.word.utils.TestFileUtil;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.STLineSpacingRule;

import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

/**
 * docx4j测试例子
 * @author liujunhui
 * @date 2020/11/10 13:13
 */
public class Docx4jSimpleExample {

    public static void main(String[] args) throws Docx4JException {
        Docx4jSimpleExample example = new Docx4jSimpleExample();
        /*File docxFile = new File(Constant.outputPath + "docx4j.docx");
        try {
            Files.createDirectories(Paths.get(docxFile.getParent()));
        } catch (IOException e) {
            e.printStackTrace();
        }

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();
        wordMLPackage.getMainDocumentPart().addParagraphOfText("hello world!");
        addStyledPText(wordMLPackage);
        wordMLPackage.save(docxFile);*/
        String path = TestFileUtil.readDeskTopPath() + "test.docx";
        // 统一设置行距
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(path));
        MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        List<Object> partContent = mainDocumentPart.getContent();
        for (Object o : partContent) {
            if (o instanceof P) {
                example.setParagraphSpacing((P) o, false, null, null, null, null,
                        true, "12", STLineSpacingRule.AT_LEAST);
            }
        }
        wordMLPackage.save(new File(path.replace("test", "new")));
    }

    /**
     * 加载读入word文件
     */
    public static void loadWord(String docxFilePath) {
        try {
            WordprocessingMLPackage template = WordprocessingMLPackage.load(new File(docxFilePath));
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

    /**
     * 设置行距
     * @param isSpace 是否设置段前段后值
     * @param before 段前磅数
     * @param after 段后磅数
     * @param beforeLines 段前行数
     * @param afterLines 段后行数
     * @param isLine 是否设置行距
     * @param lineValue 行距值
     * @param sTLineSpacingRule 自动auto 固定exact 最小 atLeast 1磅=20 1行=100 单倍行距=240
     */
    public void setParagraphSpacing(P p, boolean isSpace, String before, String after, String beforeLines, String afterLines,
                                           boolean isLine, String lineValue, STLineSpacingRule sTLineSpacingRule) {
        PPr pPr = getPPr(p);
        PPrBase.Spacing spacing = pPr.getSpacing();
        if (spacing == null) {
            spacing = new PPrBase.Spacing();
            pPr.setSpacing(spacing);
        }
        if (isSpace) {
            if (StringUtils.isNotBlank(before)) {
                // 段前磅数
                spacing.setBefore(new BigInteger(before));
            }
            if (StringUtils.isNotBlank(after)) {
                // 段后磅数
                spacing.setAfter(new BigInteger(after));
            }
            if (StringUtils.isNotBlank(beforeLines)) {
                // 段前行数
                spacing.setBeforeLines(new BigInteger(beforeLines));
            }
            if (StringUtils.isNotBlank(afterLines)) {
                // 段后行数
                spacing.setAfterLines(new BigInteger(afterLines));
            }
        }
        if (isLine) {
            if (StringUtils.isNotBlank(lineValue)) {
                spacing.setLine(new BigInteger(lineValue));
            }
            if (sTLineSpacingRule != null) {
                spacing.setLineRule(sTLineSpacingRule);
            }
        }
    }

    public PPr getPPr(P p) {
        PPr ppr = p.getPPr();
        if (ppr == null) {
            ppr = new PPr();
            p.setPPr(ppr);
        }
        return ppr;
    }

}
