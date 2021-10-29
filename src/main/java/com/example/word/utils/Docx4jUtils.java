package com.example.word.utils;

import com.aspose.words.HtmlSaveOptions;
import com.aspose.words.SaveFormat;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.docx4j.*;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.dml.CTPositiveSize2D;
import org.docx4j.dml.Graphic;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.docProps.core.CoreProperties;
import org.docx4j.docProps.extended.Properties;
import org.docx4j.finders.SectPrFinder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DocPropsCorePart;
import org.docx4j.openpackaging.parts.DocPropsCustomPart;
import org.docx4j.openpackaging.parts.DocPropsExtendedPart;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.relationships.Relationships;
import org.docx4j.vml.CTShape;
import org.docx4j.wml.*;
import org.dom4j.Attribute;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import java.io.*;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * Docx4j工具类
 * @author
 */
@Slf4j
public class Docx4jUtils {

    /**Key = docx4j.Convert.Out.HTML.OutputMethodXML*/
    public static final String DOCX4J_PARAM_04 = "docx4j.Convert.Out.HTML.OutputMethodXML";

    /**
     * @Description: 获取单元格内容,无分割符
     */
    public static String getTblContentStr(Tbl tbl) throws Exception {
        StringWriter stringWriter = new StringWriter();
        TextUtils.extractText(tbl, stringWriter);
        return stringWriter.toString();
    }

    /**
     * @Description: 获取表格内容
     */
    public static List<String> getTblContentList(Tbl tbl) throws Exception {
        List<String> resultList = new ArrayList<>();
        List<Tr> trList = getTblAllTr(tbl);
        for (Tr tr : trList) {
            StringBuilder sb = new StringBuilder();
            List<Tc> tcList = getTrAllCell(tr);
            for (Tc tc : tcList) {
                sb.append(getElementContent(tc)).append(",");
            }
            resultList.add(sb.toString());
        }
        return resultList;
    }

    /**
     * @Description: 得到表格所有的行
     */
    public static List<Tr> getTblAllTr(Tbl tbl) {
        List<Object> objList = getAllElementFromObject(tbl, Tr.class);
        List<Tr> trList = new ArrayList<>();
        if (objList == null) {
            return trList;
        }
        for (Object obj : objList) {
            if (obj instanceof Tr) {
                Tr tr = (Tr) obj;
                trList.add(tr);
            }
        }
        return trList;
    }

    /**
     * @Description: 获取所有的单元格
     */
    public static List<Tc> getTrAllCell(Tr tr) {
        List<Object> objList = getAllElementFromObject(tr, Tc.class);
        List<Tc> tcList = new ArrayList<Tc>();
        if (objList == null) {
            return tcList;
        }
        for (Object tcObj : objList) {
            if (tcObj instanceof Tc) {
                Tc objTc = (Tc) tcObj;
                tcList.add(objTc);
            }
        }
        return tcList;
    }

    /**
     * @Description:得到指定类型的元素
     */
    public static List<Object> getAllElementFromObject(Object obj,
                                                       Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
        }

        if (obj.getClass().equals(toSearch)) {
            result.add(obj);
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }

    public static String getElementContent(Object obj) throws Exception {
        StringWriter stringWriter = new StringWriter();
        TextUtils.extractText(obj, stringWriter);
        return stringWriter.toString();
    }


    public static Set<String> getFormulaStr(String wordPath) throws Docx4JException {
        Set<String> stringSet = new HashSet<>();
        WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(new File(wordPath));
        String s = XmlUtils.marshaltoString(mlPackage.getMainDocumentPart().getJaxbElement(),
                true, true);
        Matcher matcher = Pattern.compile("ProgID=\"Equation\\..*?\"").matcher(s);
        while (matcher.find()) {
            stringSet.add(matcher.group());
        }
        return stringSet;
    }


    /**
     * 对生成Document中的内容做特殊处理，去除超链接和分页符
     * @param wordMLPackage
     * @return
     */
    public static WordprocessingMLPackage deleteSpecialCharInDocument(WordprocessingMLPackage wordMLPackage) {
        try {
            // 去超链接
            String s = XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getContents(),
                    true, true);
            if (s.contains("hyperlink")) {
                s = s.replaceAll("<w:hyperlink.*?>", "")
                        .replaceAll("</w:hyperlink>", "");
            }

            // todo 测试
            if (s.contains("hlinkClick")) {
                s = s.replaceAll("<a:hlinkClick.*?>", "").replaceAll("</a:hlinkClick>", "");
            }

            // 去除图片中的r:link字段，否则出现图片无法正常打开的情况
            if (s.contains("r:link=\"rId")) {
                s = s.replaceAll("r:link=\"rId[\\d]*?\"", "");
            }
            if (s.contains("r:href=\"rId")) {
                s = s.replaceAll("r:href=\"rId[\\d]*?\"", "");
            }

            // 去分页符
            s = s.replace("<w:br w:type=\"page\"/>", "");
            Document document = (Document) XmlUtils.unmarshalString(s);
            wordMLPackage.getMainDocumentPart().setContents(document);
        } catch (Exception e) {
            e.printStackTrace();
            log.info("对document的去超链接和分页符操作失败", e);
            return null;
        }
        return wordMLPackage;
    }

    /**
     * 设置 除图片，oleobject之外的rles对应关系，而不是使用docx4j默认。
     * @param wordMLPackage
     * @param wpgNew
     */
    public static void setRelationshipsNoImageType(WordprocessingMLPackage wordMLPackage, WordprocessingMLPackage wpgNew) {

        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        // 获取模块间依赖关系
        RelationshipsPart relsPart = documentPart.getRelationshipsPart();
        Relationships rels = relsPart.getRelationships();
        List<Relationship> relList = rels.getRelationship();

        MainDocumentPart documentPartNew = wpgNew.getMainDocumentPart();
        // 获取模块间依赖关系
        RelationshipsPart relsPartNew = documentPartNew.getRelationshipsPart();
        Relationships relsNew = relsPartNew.getRelationships();
        List<Relationship> relListNew = relsNew.getRelationship();

        for (Relationship relationship : relList) {
            String type = relationship.getType();
            for (Relationship relationship1 : relListNew) {
                if (relationship1.getType().equals(type)) {
                    relationship1.setId(relationship.getId());
                }
            }
        }
    }

    /**
     * 设置 settings.xml文件 中 默认的Compat
     * @param wpgNew
     * @return
     */
    public static Boolean setDefaultCompat(WordprocessingMLPackage wpgNew) {
        //todo 这里仅添加了 行末空格显示 下划线属性，之后再需要在"高级" 中设置，在此函数中添加
        MainDocumentPart mainDocumentPart = wpgNew.getMainDocumentPart();
        DocumentSettingsPart documentSettingsPart = mainDocumentPart.getDocumentSettingsPart();
        CTCompat ctCompat  = documentSettingsPart.getJaxbElement().getCompat();
        ctCompat.setUlTrailSpace(new BooleanDefaultTrue());
        return true;
    }

    /**
     * 获取表格里的所有段落
     * @param tbl
     * @return
     */
    public static List<P> getTblAllP(Tbl tbl) {
        List<Object> objList = getAllElementFromObject(tbl, P.class);
        List<P> pList = new ArrayList<>();
        if (objList == null) {
            return pList;
        }
        for (Object obj : objList) {
            if (obj instanceof P) {
                P p = (P) obj;
                pList.add(p);
            }
        }
        return pList;
    }


    /**
     * 获取run中t的文本
     *
     * @param
     * @return
     */
    public static String getText(R run) {
        List children = run.getContent();
        RPr rPr = run.getRPr();
        StringBuilder runText = new StringBuilder();
        JAXBElement<Text> textJAXBElement = new ObjectFactory().createRT(new Text());
        for (Object o : children) {
            if (o instanceof JAXBElement) {
                if (textJAXBElement.getName().equals(((JAXBElement) o).getName())) {
                    Text t = (Text) ((JAXBElement) o).getValue();
                    // 下划线
                    if (rPr != null && rPr.getU() != null && StringUtils.isNotEmpty(t.getValue()) && StringUtils.isWhitespace(t.getValue())) {
                        for (int i = 0; i < t.getValue().length(); i++) {
                            runText.append("_");
                        }
                        t.setValue(runText.toString());
                        //防止出现双下划线
                        rPr.setU(null);
                    } else {
                        runText.append(t.getValue());
                    }
                }
            }
        }
        return runText.toString();
    }

    /**
     * 获取P中文本，包括R, Smart, 文本框中的文本
     * @param p
     * @return
     */
    public static String getText(P p) {

        List<Object> rContent = p.getContent();
        StringBuilder runText = new StringBuilder();
        for (Object run : rContent) {
            if (run instanceof R) {
                runText.append(getText((R) run));
            }
            if (run instanceof JAXBElement) {
                JAXBElement<CTSmartTagRun> pSmartTag = new ObjectFactory().createPSmartTag(new CTSmartTagRun());
                if (pSmartTag.getName().equals(((JAXBElement) run).getName())) {
                    CTSmartTagRun smartTagRun = (CTSmartTagRun) ((JAXBElement) run).getValue();
                    List<Object> smartTagRunContent = smartTagRun.getContent();
                    for (Object r : smartTagRunContent) {
                        if (r instanceof R) {
                            runText.append(getText((R) r));
                        }
                    }
                }
            }
        }

        // 通过xml的方式直接解析文本
        String s = XmlUtils.marshaltoString(p);
        // 如果同时存在，不需要fallback标签
        if (s.contains("mc:Fallback") && s.contains("mc:Choice")) {
            Matcher matcher = Pattern.compile("<mc:Fallback.*?>.*?</mc:Fallback>", Pattern.DOTALL).matcher(s);
            if (matcher.find()) {
                s = s.replace(matcher.group(), "");
            }
        }

        // 获取文本框中的文本
        Matcher matcher = Pattern.compile("<mc:Choice.*?>.*?</mc:Choice>", Pattern.DOTALL).matcher(s);
        while (matcher.find()) {
            String txText = matcher.group();
            // 匹配文本标签
            Matcher matcher1 = Pattern.compile("<w:t( xml:space=\".*?\")?>(.*?)</w:t>", Pattern.DOTALL).matcher(txText);
            while (matcher1.find()) {
                String group = matcher1.group(2);
                runText.append(group);
            }
        }
        return runText.toString();
    }

    /**
     * word转html
     * @param docxFile
     * @param outFile
     * @return
     */
    public static boolean writeToHtmlBak(File docxFile, File outFile) {
        OutputStream output = null;
        WordprocessingMLPackage wmlPackage = null;
        try {
            wmlPackage = WordprocessingMLPackage.load(docxFile);
            setNumberingDefinitionsPart(wmlPackage);
            removeLabel(wmlPackage, "w:pStyle");
            // 保存图片的文件夹
            String image =  FilenameUtils.getBaseName(outFile.getAbsolutePath()) + ".files";
            String imageFilePath = outFile.getParent() + "/" + image + "/";
            if (!new File(imageFilePath).exists()) {
                new File(imageFilePath).mkdirs();
            }
            //创建文件输出流
            output = new FileOutputStream(outFile);
            //创建Html输出设置
            HTMLSettings htmlSettings = Docx4J.createHTMLSettings();
            htmlSettings.setImageDirPath(imageFilePath);
            htmlSettings.setImageTargetUri(image);
            htmlSettings.setWmlPackage(wmlPackage);
            Docx4jProperties.setProperty(DOCX4J_PARAM_04, true);
            Docx4J.toHTML(htmlSettings, output, Docx4J.FLAG_EXPORT_PREFER_XSL);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            IOUtils.closeQuietly(output);
        }
        return true;
    }

    public static boolean writeToHtml(File docxFile, File outFile) {
        log.info("正在转换html...");
        // 保存图片的文件夹
        String image =  FilenameUtils.getBaseName(outFile.getAbsolutePath()) + ".files";
        String imageFilePath = outFile.getParent() + "/" + image + "/";
        if (!new File(imageFilePath).exists()) {
            new File(imageFilePath).mkdirs();
        }
        try {
            com.aspose.words.Document document = new com.aspose.words.Document(docxFile.getAbsolutePath());
            HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions(SaveFormat.HTML);
            htmlSaveOptions.setImagesFolder(imageFilePath);
            document.save(outFile.getAbsolutePath(), htmlSaveOptions);
        } catch (Exception e) {
            log.error("转换html失败：", e);
            return false;
        }
        return true;
    }

    public static void removeLabel(WordprocessingMLPackage wmlPackage, String xmlLabel) {
        if (wmlPackage == null) {
            return;
        }
        if (StringUtils.isBlank(xmlLabel)) {
            return;
        }
        MainDocumentPart mdp = wmlPackage.getMainDocumentPart();
        String s = XmlUtils.marshaltoString(mdp.getJaxbElement());
        org.dom4j.Document document = getDocument(s);
        List<Element> pStyleList = getElementByQualifiedName(document.getRootElement(), xmlLabel);
        for (Element element : pStyleList) {
            element.getParent().remove(element);
        }
        Document document1 = null;
        try {
            String xml = document.asXML();
            document1 = (Document) XmlUtils.unmarshalString(xml);
        } catch (JAXBException e) {
            log.error("反序列化失败", e);
            return;
        }
        mdp.setJaxbElement(document1);
    }

    /**
     * 删除学科网以及菁优网信息
     * @param sourcePath word绝对路径
     * @param targetPath 目标绝对路径
     */
    public static boolean removeZxxkOrJyeooInfo(String sourcePath, String targetPath) {
        File file = new File(sourcePath);
        if (!file.exists()) {
            return false;
        }
        WordprocessingMLPackage mlPackage = null;
        try {
            mlPackage = WordprocessingMLPackage.load(file);
            Docx4jUtils.removeZxxkOrJyeooInfo(mlPackage);
            File file1 = new File(targetPath);
            if (!file1.exists()) {
                FileUtils.touch(file1);
            }
            mlPackage.save(file1);
        } catch (Exception e) {
            log.error("", e);
            return false;
        }
        return true;
    }

    /**
     * 删除学科网以及菁优网信息
     * @param bytes z
     */
    public static byte[] removeZxxkOrJyeooInfo(byte[] bytes) {
        if (bytes == null) {
            return null;
        }
        WordprocessingMLPackage mlPackage = null;
        try {
            mlPackage = WordprocessingMLPackage.load(new ByteArrayInputStream(bytes));
            Docx4jUtils.removeZxxkOrJyeooInfo(mlPackage);
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            mlPackage.save(outputStream);
            return outputStream.toByteArray();
        } catch (Exception e) {
            log.error("", e);
            return null;
        }
    }

    /**
     * 删除学科网以及菁优网信息
     * @param wordMLPackage
     */
    public static void removeZxxkOrJyeooInfo(WordprocessingMLPackage wordMLPackage) {
        // 删除页眉页脚
        removeHeaderAndFooter(wordMLPackage);
        // 删除字体小于10磅，或字体颜色为白色的文本
        removeSmallTextOrWhiteText(wordMLPackage);
        // 删除属性：descr，alt
        removeDescrAndAltAttribute(wordMLPackage);
        // 删除文档属性
        removeDocProps(wordMLPackage);
    }

    /**
     * 删除文档属性
     * @param wordMLPackage
     */
    private static void removeDocProps(WordprocessingMLPackage wordMLPackage) {
        DocPropsCorePart docPropsCorePart = wordMLPackage.getDocPropsCorePart();
        DocPropsExtendedPart docPropsExtendedPart = wordMLPackage.getDocPropsExtendedPart();
        DocPropsCustomPart docPropsCustomPart = wordMLPackage.getDocPropsCustomPart();
        if (docPropsCorePart != null) {
            CoreProperties coreProperties = docPropsCorePart.getJaxbElement();
            coreProperties.setCategory(null);
            coreProperties.setContentStatus(null);
            coreProperties.setContentType(null);
            // 创建者
            coreProperties.setCreator(null);
            coreProperties.setDescription(null);
            coreProperties.setIdentifier(null);
            // 标记
            coreProperties.setKeywords(null);
            // 上次修改者
            coreProperties.setLastModifiedBy(null);
        }
        if (docPropsExtendedPart != null) {
            Properties properties = docPropsExtendedPart.getJaxbElement();
            properties.setManager(null);
            properties.setCompany(null);
            properties.setHyperlinkBase(null);
        }
        if (docPropsCustomPart != null) {
            org.docx4j.docProps.custom.Properties properties = docPropsCustomPart.getJaxbElement();
            properties.getProperty().clear();
        }
    }

    /**
     * 删除页眉页脚
     * @param wordMLPackage
     */
    public static void removeHeaderAndFooter(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
        // Remove from sectPr
        SectPrFinder finder = new SectPrFinder(mdp);
        new TraversalUtil(mdp.getContent(), finder);
        for (SectPr sectPr : finder.getOrderedSectPrList()) {
            sectPr.getEGHdrFtrReferences().clear();
        }
        // Remove rels
        List<Relationship> hfRels = new ArrayList<>();
        for (Relationship rel : mdp.getRelationshipsPart().getRelationships().getRelationship() ) {
            if (rel.getType().equals(Namespaces.HEADER) || rel.getType().equals(Namespaces.FOOTER)) {
                hfRels.add(rel);
            }
        }
        for (Relationship rel : hfRels ) {
            mdp.getRelationshipsPart().removeRelationship(rel);
        }
    }

    /**
     * 删除字体比较小的文本以及白颜色的文本
     * 删除r标签内的，对于其他标签内的不予删除
     *
     * @param wordMLPackage
     */
    public static void removeSmallTextOrWhiteText(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
        String s = XmlUtils.marshaltoString(mdp.getJaxbElement());
        //创建解析器对象
        SAXReader saxReader = new SAXReader();
        //得到document
        org.dom4j.Document document = null;
        try {
            document = saxReader.read(new ByteArrayInputStream(s.getBytes()));
        } catch (DocumentException e) {
            log.error("removeSmallTextOrWhiteText:", e);
            return;
        }
        //得到根节点
        org.dom4j.Element root = document.getRootElement();
        // 递归遍历，获取所有的color对象（dom4j和docx4j一样，修改的是对象的引用）
        // 获取颜色信息
        List<Element> elements = getObjectElement(root, "color");
        for (Element element : elements) {
            String value = element.attributeValue("val");
            // 值为白色的
            if (StringUtils.equalsIgnoreCase(value, "FFFFFF")) {
                Element parent = element.getParent();
                if (parent != null && parent.getName().equals("rPr")) {
                    removeElement(parent);
                }
            }
        }
        List<Element> elements1 = getObjectElement(root, "sz");
        for (Element element : elements1) {
            String value = element.attributeValue("val");
            // 字体大小小于5磅的
            if (StringUtils.isNumeric(value) && Integer.valueOf(value) < 10) {
                Element parent = element.getParent();
                if (parent != null && parent.getName().equals("rPr")) {
                    Element szCs = parent.element("szCs");
                    if (szCs != null) {
                        String szCsValue = szCs.attributeValue("val");
                        // 如果sz和szCs标签两者字体大小相等，则进行删除
                        if (value.equals(szCsValue)) {
                            removeElement(parent);
                        }
                    }
                }
            }
        }
        Document document1 = null;
        try {
            String xml = document.asXML();
            document1 = (Document) XmlUtils.unmarshalString(xml);
        } catch (JAXBException e) {
            log.error("反序列化失败", e);
            return;
        }
        mdp.setJaxbElement(document1);
    }

    private static void removeElement(Element parent) {
        Element parentParent = parent.getParent();
        if (parentParent.getName().equals("r")) {
            List<Element> t = getObjectElement(parentParent, "t");
            String text = t.stream().map(Element::getTextTrim).collect(Collectors.joining());
            log.info("删除的文本为：{}", text);
            parentParent.getParent().remove(parentParent);
        }
        if (parentParent.getName().equalsIgnoreCase("pPr")) {
            parentParent.getParent().remove(parentParent);
        }
    }

    private static List<Element> getObjectElement(Element root, String name) {
        List elements = root.elements();
        List<Element> elementList = new ArrayList<>();
        for (Object element : elements) {
            if (element instanceof Element) {
                if (((Element) element).getName().equals(name)) {
                    elementList.add((Element) element);
                } else {
                    elementList.addAll(getObjectElement((Element) element, name));
                }
            }
        }
        return elementList;
    }

    public static List<Element> getElementByQualifiedName(Element root, String name) {
        List elements = root.elements();
        List<Element> elementList = new ArrayList<>();
        for (Object element : elements) {
            if (element instanceof Element) {
                if (((Element) element).getQualifiedName().equals(name)) {
                    elementList.add((Element) element);
                } else {
                    elementList.addAll(getElementByQualifiedName((Element) element, name));
                }
            }
        }
        return elementList;
    }

    /**
     * 删除alt，以及descr属性的值
     * 因为这些属性会包含学科网，菁优网信息，移除之
     * @param wordMLPackage
     */
    public static void removeDescrAndAltAttribute(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();
        String s = XmlUtils.marshaltoString(mdp.getJaxbElement());
        //创建解析器对象
        SAXReader saxReader = new SAXReader();
        //得到document
        org.dom4j.Document document = null;
        try {
            document = saxReader.read(new ByteArrayInputStream(s.getBytes()));
        } catch (DocumentException e) {
            log.error("removeSmallTextOrWhiteText:", e);
            return;
        }
        Element rootElement = document.getRootElement();
        List<Attribute> objectAttribute = getObjectAttribute(rootElement, "descr", "alt");
        for (Attribute attribute : objectAttribute) {
            String value = attribute.getValue();
            if (StringUtils.isNotBlank(value)) {
                log.debug("移除的属性值：{}:{}", attribute.getName(), value);
                attribute.getParent().remove(attribute);
            }
        }
        Document document1 = null;
        try {
            String xml = document.asXML();
            document1 = (Document) XmlUtils.unmarshalString(xml);
        } catch (JAXBException e) {
            log.error("反序列化失败", e);
            return;
        }
        mdp.setJaxbElement(document1);
    }

    private static List<Attribute> getObjectAttribute(Element root, String... attributeNames) {
        List elements = root.elements();
        List<Attribute> attributeList = new ArrayList<>();
        if (elements.size() == 0) {
            return attributeList;
        }
        for (Object element : elements) {
            if (element instanceof Element) {
                for (String name : attributeNames) {
                    Attribute attribute = ((Element) element).attribute(name);
                    if (attribute != null) {
                        attributeList.add(attribute);
                    }
                    List<Attribute> objectAttribute = getObjectAttribute((Element) element, name);
                    attributeList.addAll(objectAttribute);
                }
            }
        }
        return attributeList;
    }

    /**
     * 获取dom4j的文档句柄，每次都写这一段话很烦
     *
     * @param s 需要解析的内容
     * @return 文档对象或null
     */
    public static org.dom4j.Document getDocument(String s) {
        //创建解析器对象
        SAXReader saxReader = new SAXReader();
        //得到document
        org.dom4j.Document document = null;
        try {
            document = saxReader.read(new ByteArrayInputStream(s.getBytes()));
        } catch (DocumentException e) {
            log.error("removeSmallTextOrWhiteText:", e);
        }
        return document;
    }

    /**
     * 把软换行变为硬换行
     *
     * @param p
     * @return
     */
    private static List<P> break2Para(P p) {
        List<P> pList = new ArrayList<>();
        if (p == null) {
            pList.add(p);
            return pList;
        }
        String s = XmlUtils.marshaltoString(p);
        org.dom4j.Document document = getDocument(s);
        List<Element> brElement = getObjectElement(document.getRootElement(), "br");
        if (CollectionUtils.isEmpty(brElement)) {
            pList.add(p);
            return pList;
        }
        ObjectFactory factory = new ObjectFactory();
        PPr pPr = p.getPPr();
        if (pPr == null) {
            pPr = factory.createPPr();
        }
        PPr pPr1 = XmlUtils.deepCopy(pPr);
        if (pPr1.getNumPr() != null) {
            pPr1.setNumPr(null);
        }

        List<Object> content = p.getContent();
        // 目前只支持p->r两级
        P pPrep = factory.createP();
        pPrep.setPPr(pPr);
        for (Object o : content) {
            if (o instanceof R) {
                RPr rPr = ((R) o).getRPr();
                if (rPr == null) {
                    rPr = factory.createRPr();
                }
                R rPrep = factory.createR();
                rPrep.setRPr(XmlUtils.deepCopy(rPr));
                pPrep.getContent().add(rPrep);
                List<Object> content1 = ((R) o).getContent();
                for (Object o1 : content1) {
                    if (o1 instanceof Br) {
                        pList.add(pPrep);
                        pPrep = factory.createP();
                        pPrep.setPPr(pPr1);
                        rPrep = factory.createR();
                        rPrep.setRPr(XmlUtils.deepCopy(rPr));
                        pPrep.getContent().add(rPrep);
                    } else {
                        rPrep.getContent().add(o1);
                    }
                }
            } else {
                pPrep.getContent().add(o);
            }
        }
        pList.add(pPrep);
        return pList;
    }

    /**
     * 把软换行变为硬换行
     *
     * @param mlPackage
     */
    public static void break2Para(WordprocessingMLPackage mlPackage) {
        List<Object> content = mlPackage.getMainDocumentPart().getContent();
        List<Object> h = new ArrayList<>();
        for (Object o : content) {
            if (o instanceof P) {
                List<P> pList = break2Para((P) o);
                h.addAll(pList);
            } else {
                h.add(o);
            }
        }
        content.clear();
        content.addAll(h);
    }

    public static boolean setNumberingDefinitionsPart(WordprocessingMLPackage mlPackage) {
        NumberingDefinitionsPart numberingDefinitionsPart = mlPackage.getMainDocumentPart().getNumberingDefinitionsPart();

        // 获取存在的编号信息
        if (numberingDefinitionsPart == null) {
            // 若为null，需要初始化
            NumberingDefinitionsPart ndp = null;
            try {
                ndp = new NumberingDefinitionsPart();
                mlPackage.getMainDocumentPart().addTargetPart(ndp);
            } catch (InvalidFormatException e) {
                log.error("非法格式异常：", e);
                return false;
            }
            final String initialNumbering = "<w:numbering xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
                    + "</w:numbering>";
            try {
                ndp.setJaxbElement((Numbering) XmlUtils.unmarshalString(initialNumbering));
            } catch (JAXBException e) {
                e.printStackTrace();
                return false;
            }
        }
        return true;
    }

    /**
     * 如果text为空，且存在分节符，去除 存在的 相应的编号
     * @param p
     */
    public static void deleteNumPrIfNoExistText(WordprocessingMLPackage mlPackage) {

        List<Object> content = mlPackage.getMainDocumentPart().getContent();
        for (Object o : content) {
            if (o instanceof P) {
                P p = (P) o;
                PPr pPr = p.getPPr();
                if (pPr == null || pPr.getNumPr() == null) {
                    continue;
                }
                String text = Docx4jUtils.getText(p);
                if (StringUtils.isEmpty(text) && pPr.getSectPr() != null) {
                    pPr.setNumPr(null);
                }
            }
        }
    }

    /**
     * 删除小图片
     * 如果图片像素小于5，则删除
     * pt=px乘以3/4。
     * @param mlPackage
     */
    public static void deleteSmallImg(WordprocessingMLPackage mlPackage) {

        Integer maxSize = 5;

        List<Object> content = mlPackage.getMainDocumentPart().getContent();
        for (Object o : content) {
            List<Object> drawingObjectList = Docx4jUtils.getAllElementFromObject(o, Drawing.class);
            for (Object drawingObject : drawingObjectList) {
                List<Object> anchorOrInline = ((Drawing) drawingObject).getAnchorOrInline();
                for (Object anchorOrInlineObject : anchorOrInline) {
                    CTPositiveSize2D extent = null;
                    Graphic graphic = null;
                    if (anchorOrInlineObject instanceof Inline) {
                        extent = ((Inline) anchorOrInlineObject).getExtent();
                        graphic = ((Inline) anchorOrInlineObject).getGraphic();
                    }
                    if (anchorOrInlineObject instanceof Anchor) {
                        extent = ((Anchor) anchorOrInlineObject).getExtent();
                        graphic = ((Anchor) anchorOrInlineObject).getGraphic();
                    }
                    if (extent != null) {
                        long cx = extent.getCx();
                        long cy = extent.getCy();
                        deleteSmallImgGraphic(drawingObject, cx, cy, maxSize);
                    }
                    if (graphic != null && graphic.getGraphicData() != null
                            && graphic.getGraphicData().getPic() != null
                            && graphic.getGraphicData().getPic().getSpPr() != null
                            && graphic.getGraphicData().getPic().getSpPr().getXfrm() != null) {
                        CTPositiveSize2D ext = graphic.getGraphicData().getPic().getSpPr().getXfrm().getExt();
                        long cx = ext.getCx();
                        long cy = ext.getCy();
                        deleteSmallImgGraphic(drawingObject, cx, cy, maxSize);
                    }
                }
            }
            List<Object> objectList = Docx4jUtils.getAllElementFromObject(o, CTObject.class);
            for (Object object : objectList) {
                List<Object> anyAndAny = ((CTObject) object).getAnyAndAny();
                for (Object o1 : anyAndAny) {
                    List<Object> allElementFromObject = Docx4jUtils.getAllElementFromObject(o1, CTShape.class);
                    for (Object o2 : allElementFromObject) {
                        String style = ((CTShape) o2).getStyle();
                        if (style != null && (style.contains("height") || style.contains("width"))) {
                            Matcher widthMatcher = Pattern.compile("width:(.*?)pt").matcher(style);
                            Matcher heightMatcher = Pattern.compile("height:(.*?)pt").matcher(style);
                            deleteSmallImgObject(maxSize, object, widthMatcher);
                            deleteSmallImgObject(maxSize, object, heightMatcher);
                        }
                    }
                }
            }
        }
    }

    private static void deleteSmallImgGraphic(Object drawingObject, long cx, long cy, Integer maxSize) {
        long twipToEMU = UnitsOfMeasurement.twipToEMU(UnitsOfMeasurement.pointToTwip(maxSize));
        if (cx < twipToEMU || cy < twipToEMU) {
            Object parent = ((Drawing) drawingObject).getParent();
            if (parent instanceof R) {
                log.info("删除文档中像素小于5的图片");
                ((R) parent).getContent().remove(getJAXBElementByValue((R) parent, drawingObject));
            }
        }
    }

    private static void deleteSmallImgObject(Integer maxSize, Object object, Matcher matcher) {
        if (matcher.find()) {
            String width = matcher.group(1);
            if (StringUtils.isBlank(width) || !NumberUtils.isParsable(width)) {
                return;
            }
            if (Double.valueOf(width).compareTo(Double.valueOf(maxSize)) < 0) {
                Object parent = ((CTObject) object).getParent();
                if (parent instanceof R) {
                    log.info("删除文档中像素小于5的图片");
                    ((R) parent).getContent().remove(getJAXBElementByValue((R) parent, object));
                }
            }
        }
    }

    private static JAXBElement getJAXBElementByValue(R run, Object value) {
        List<Object> content = run.getContent();
        for (Object o : content) {
            if (o instanceof JAXBElement) {
                if (((JAXBElement) o).getValue() == value) {
                    return (JAXBElement) o;
                }
            }
        }
        return null;
    }
}
