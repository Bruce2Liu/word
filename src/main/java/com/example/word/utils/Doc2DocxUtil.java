package com.example.word.utils;

import com.aspose.words.Document;
import com.aspose.words.License;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.InputStream;

/**
 * @author liujunhui
 * @date 2021/1/9 16:27
 */
public class Doc2DocxUtil {

    /**
     * doc文档转化docx格式文档
     * @param filePath    原文件绝对路径
     * @param targetPath  生成的文件的绝对路径
     * @return
     */
    public static boolean doc2Docx(String filePath, String targetPath) {
        if (!filePath.endsWith(".doc")) {
            System.out.println("doc2Docx方法待转化文件不是doc文件");
            return false;
        }
        if (!getLicense()) {
            System.out.println("证书加载失败！");
            return false;
        }

        try {
            Document doc = new Document(filePath);
            doc.removeMacros();
            doc.save(targetPath);
            if (!isNormal(new File(targetPath))) {
                return false;
            }
        } catch (Exception e) {
            return false;
        }
        return true;
    }

    /**
     * 获取license
     * 加载证书,解决水印问题,文档转化过程中的截断问题等等
     * @return
     */
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = Doc2DocxUtil.class.getClassLoader().getResourceAsStream("aspose-word/license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {

        }
        return result;
    }

    /**
     * 检查doc转换后的文件是否正常
     * @param file
     * @return
     * @throws Docx4JException
     */
    public static boolean isNormal(File file) throws Docx4JException {
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(file);
        String s = XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getContents(),
                true, true);
        String error = "This document was truncated here because it was created using Aspose.Words in Evaluation Mode";
        return !s.contains(error);
    }

    public static void main(String[] args) {
        String deskFile = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
        System.out.println(deskFile);
    }
}
