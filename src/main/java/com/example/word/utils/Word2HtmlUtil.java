package com.example.word.utils;

import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.IImageSavingCallback;
import com.aspose.words.ImageSavingArgs;
import com.aspose.words.SaveFormat;
import lombok.extern.slf4j.Slf4j;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class Word2HtmlUtil {

    /** word中的符号，以及替代物 */
    private static Map<String, String> wordMap = new HashMap<>();
    /** 替代物 到 html 的映射，都是正则表达式 */
    private static Map<String, String> htmlMap = new HashMap<>();

    static {
        // word中的着重号，xml 标识：<w:em w:val="dot"/>
        // 替代品：⌊dot⌉-⌊/dot⌉，其中横线为起始的分割线
        wordMap.put("dot", "⌊dot⌉-⌊/dot⌉");
        // 着重号在 html 中的标签，沿用语文html试卷中的写法
        htmlMap.put("⌊dot⌉(.*?)⌊/dot⌉", "<spec class=\"point\" style=\"border-bottom: 3px dotted black\">$1</spec>");
    }

    public static boolean writeToHtml(InputStream docxFile, OutputStream outFile, Map<String, byte[]> imgMap) {
        log.info("正在转换html...");

        ByteArrayOutputStream prepWord = new ByteArrayOutputStream();
        // 对原始的IO流进行备份
        boolean prepWriteHtml = prepWriteHtml(docxFile, prepWord);
        if (!prepWriteHtml) {
            // 如果失败，
            log.error("处理html特殊标签失败");
            return false;
        }
        Map<String, ByteArrayOutputStream> map = new LinkedHashMap<>();
        try {
            com.aspose.words.Document document = new com.aspose.words.Document(new ByteArrayInputStream(prepWord.toByteArray()));
            HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions(SaveFormat.HTML);
            htmlSaveOptions.setImageSavingCallback(
                    new IImageSavingCallback() {
                        @Override
                        public void imageSaving(ImageSavingArgs imageSavingArgs) throws Exception {
                            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                            imageSavingArgs.setImageStream(outputStream);
                            map.put(imageSavingArgs.getImageFileName(), outputStream);
                        }
                    }
            );
            // 此处执行 setImageSavingCallback
            document.save(outFile, htmlSaveOptions);
        } catch (Exception e) {
            log.error("转换html失败：", e);
            return false;
        }
        // 当word中有图片，但是传过来的map为null，则返回false
        if (!map.isEmpty() && imgMap == null) {
            log.error("imgMap此时不能为null");
            return false;
        }
        for (Map.Entry<String, ByteArrayOutputStream> stringByteArrayOutputStreamEntry : map.entrySet()) {
            String key = stringByteArrayOutputStreamEntry.getKey();
            ByteArrayOutputStream value = stringByteArrayOutputStreamEntry.getValue();
            imgMap.put(key, value.toByteArray());
        }
        return true;
    }

    private static boolean prepWriteHtml(InputStream docxFile, OutputStream outFile) {
        log.info("正在进行html标签替换");
        WordprocessingMLPackage mlPackage = null;
        try {
            mlPackage = WordprocessingMLPackage.load(docxFile);
        } catch (Docx4JException e) {
            log.error("", e);
            return false;
        }
        List<Object> content = mlPackage.getMainDocumentPart().getContent();
        for (Object o : content) {
            if (o instanceof P) {
                for (String key : wordMap.keySet()) {
                    if (key.equals("dot")) {
                        replaceDot(o, key);
                    }
                }
            }
            if (o instanceof JAXBElement) {
                Object oClass = ((JAXBElement) o).getValue();
                if (oClass.getClass().equals(Tbl.class)) {
                    Tbl tbl = (Tbl) oClass;
                    List<P> tblAllP = Docx4jUtils.getTblAllP(tbl);
                    for (P p : tblAllP) {
                        for (String key : wordMap.keySet()) {
                            if (key.equals("dot")) {
                                replaceDot(p, key);
                            }
                        }
                    }
                }
            }
        }
        try {
            mlPackage.save(outFile);
        } catch (Docx4JException e) {
            log.error("", e);
            return false;
        }
        return true;
    }

    private static void replaceDot(Object o, String key) {
        boolean isDot = false;
        PPr pPr = ((P) o).getPPr();
        if (pPr != null) {
            ParaRPr rPr = pPr.getRPr();
            if (rPr != null) {
                CTEm em = rPr.getEm();
                // 着重号
                if (em != null && em.getVal() == STEm.DOT) {
                    isDot = true;
                }
            }
        }
        List<Object> objects = ((P) o).getContent();
        for (Object object : objects) {
            if (object instanceof R) {
                boolean isRunDot = false;
                RPr rPr = ((R) object).getRPr();
                if (rPr != null) {
                    CTEm em = rPr.getEm();
                    // 着重号
                    if (em != null && em.getVal() == STEm.DOT) {
                        isRunDot = true;
                    }
                }
                if (isDot || isRunDot) {
                    List<Object> textObject = ((R) object).getContent();
                    // 替换着重号
                    for (Object text : textObject) {
                        JAXBElement<Text> textJAXBElement = new ObjectFactory().createRT(new Text());
                        if (text instanceof JAXBElement) {
                            if (textJAXBElement.getName().equals(((JAXBElement) text).getName())) {
                                Text t = (Text) ((JAXBElement) text).getValue();
                                String dot = wordMap.get(key);
                                String value = t.getValue();
                                String[] split = dot.split("-");
                                t.setValue(split[0] + value + split[1]);
                            }
                        }
                    }
                }
            }
        }
    }

    public static String postWriteHtml(String htmlFile) {
        for (Map.Entry<String, String> stringStringEntry : htmlMap.entrySet()) {
            String key = stringStringEntry.getKey();
            String value = stringStringEntry.getValue();
            htmlFile = htmlFile.replaceAll(key, value);
        }
        return htmlFile;
    }
}
