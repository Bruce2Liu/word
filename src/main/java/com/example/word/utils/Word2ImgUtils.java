package com.example.word.utils;

import com.aspose.words.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.util.Base64Utils;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class Word2ImgUtils {

    /**
     * 返回一张图片，字节流的形式
     *
     * @param inputStream
     * @return
     * @throws Exception
     */
    public static byte[] word2Img(InputStream inputStream) throws Exception {

        log.info("======word2ImgOne======");
        long l = System.currentTimeMillis();
        Document document = new Document(inputStream);

        document.removeMacros();
        document.removeSmartTags();
        document.removeExternalSchemaReferences();

        removeHeadersFooters(document);

        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.save(outputStream, options);

        ByteArrayOutputStream outputStream1 = new ByteArrayOutputStream();
        ImageCutUtils.cut(outputStream.toByteArray(), outputStream1);
        byte[] bytes = outputStream1.toByteArray();

        outputStream.close();
        outputStream1.close();
        log.info("本次图片转化耗时：{}s", (System.currentTimeMillis() - l) / 1000.0);
        return bytes;
    }

    /**
     * 返回多张图片，base64的形式
     *
     * @param inputStream
     * @param isCut 是否需要剪切图片
     * @return
     * @throws Exception
     */
    public static List<String> word2ImgBase64List(InputStream inputStream, boolean isCut) throws Exception {

        log.info("进入方法：word2Img");
        long l = System.currentTimeMillis();
        Document document = new Document(inputStream);

        document.removeMacros();
        document.removeSmartTags();
        document.removeExternalSchemaReferences();

        removeHeadersFooters(document);

        List<String> imgBase64List = new ArrayList<>();
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.PNG);
        int pageCount = document.getPageCount();
        for (int i = 0; i < pageCount; i++) {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            options.setPageIndex(i);
            document.save(outputStream, options);

            if (isCut) {
                ByteArrayOutputStream outputStream1 = new ByteArrayOutputStream();
                ImageCutUtils.cut(outputStream.toByteArray(), outputStream1);
                imgBase64List.add(Base64Utils.encodeToString(outputStream1.toByteArray()));
                outputStream1.close();
            } else {
                imgBase64List.add(Base64Utils.encodeToString(outputStream.toByteArray()));
            }
            outputStream.close();
        }
        log.info("本次图片转化耗时：{}s", (System.currentTimeMillis() - l) / 1000.0);
        return imgBase64List;
    }

    private static void removeHeadersFooters(Document document) {
        for (Section section : document.getSections()) {
            // Up to three different footers are possible in a section (for first, even and odd pages).
            // We check and delete all of them.
            HeaderFooter footer;
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_FIRST);
            if (footer != null) {
                footer.remove();
            }
            // Primary footer is the footer used for odd pages.
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY);
            if (footer != null) {
                footer.remove();
            }
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_EVEN);
            if (footer != null) {
                footer.remove();
            }
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_EVEN);
            if (footer != null) {
                footer.remove();
            }
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_FIRST);
            if (footer != null) {
                footer.remove();
            }
            footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY);
            if (footer != null) {
                footer.remove();
            }
        }
        // 设置背景图片为null
        document.getDocument().setBackgroundShape(null);
    }
}
