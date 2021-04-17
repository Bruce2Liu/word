package com.example.word.utils;

import org.apache.commons.collections4.CollectionUtils;

import javax.imageio.ImageIO;
import javax.imageio.ImageReadParam;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;

public class ImageCutUtils {
    private static int getInitWidth(BufferedImage img) {
        int height = img.getHeight();
        int width = img.getWidth();
        int initWidth = 0;
        java.util.List<Integer> initWidthList = new ArrayList<>();
        for (int i = 0; i < height; i++) {
            for (int j = 0; j < width; j++) {
                if (img.getRGB(j, i) != Color.WHITE.getRGB()) {
                    initWidth = j;
                    initWidthList.add(initWidth);
                    break;
                }
            }
        }
        initWidthList.sort(Comparator.comparingInt(x -> x));
        if (CollectionUtils.isNotEmpty(initWidthList)) {
            initWidth = initWidthList.get(0) - 1;
        }
        return initWidth;
    }

    private static int getInitHeight(BufferedImage img) {
        int width = img.getWidth();
        int height = img.getHeight();
        java.util.List<Integer> initHeightList = new ArrayList<>();
        int initHeight = 0;

        for (int i = 0; i < width; i++) {
            for (int j = 0; j < height; j++) {
                if (img.getRGB(i, j) != Color.WHITE.getRGB()) {
                    initHeight = j;
                    initHeightList.add(initHeight);
                    break;
                }
            }
        }
        initHeightList.sort(Comparator.comparingInt(x -> x));
        if (CollectionUtils.isNotEmpty(initHeightList)) {
            initHeight = initHeightList.get(0) - 1;
        }
        return initHeight;
    }

    private static int getTrimmedWidth(BufferedImage img) {
        int height = img.getHeight();
        int width = img.getWidth();
        int trimmedWidth = width;
        java.util.List<Integer> trimmedWidthList = new ArrayList<>();
        for (int i = 0; i < height; i++) {
            for (int j = width - 1; j >= 0; j--) {
                if (img.getRGB(j, i) != Color.WHITE.getRGB()) {
                    trimmedWidth = j;
                    trimmedWidthList.add(trimmedWidth);
                    break;
                }
            }
        }
        trimmedWidthList.sort(Comparator.comparingInt(x -> x));
        if (CollectionUtils.isNotEmpty(trimmedWidthList)) {
            trimmedWidth = trimmedWidthList.get(trimmedWidthList.size() - 1) + 1;
        }
        return trimmedWidth;
    }

    private static int getTrimmedHeight(BufferedImage img) {
        int width = img.getWidth();
        int height = img.getHeight();
        java.util.List<Integer> trimmedHeightList = new ArrayList<>();
        int trimmedHeight = height;

        for (int i = 0; i < width; i++) {
            for (int j = height - 1; j >= 0; j--) {
                if (img.getRGB(i, j) != Color.WHITE.getRGB()) {
                    trimmedHeight = j;
                    trimmedHeightList.add(trimmedHeight);
                    break;
                }
            }
        }
        trimmedHeightList.sort(Comparator.comparingInt(x -> x));
        if (CollectionUtils.isNotEmpty(trimmedHeightList)) {
            trimmedHeight = trimmedHeightList.get(trimmedHeightList.size() - 1) + 1;
        }
        return trimmedHeight;
    }

    /**
     * 图片裁剪
     *
     * @param src 原始图片输入流
     * @param out 成品图片输出流
     * @throws IOException
     */
    public static void cut(byte[] src, OutputStream out) throws IOException {
        BufferedImage img = ImageIO.read(new ByteArrayInputStream(src));
        int x = getInitWidth(img);
        int y = getInitHeight(img);
        int width = getTrimmedWidth(img) - x;
        int height = getTrimmedHeight(img) - y;
        cut(new ByteArrayInputStream(src), out, x, y, width, height);
    }

    /**
     * 图片裁剪
     *
     * @param src    原始图片输入流
     * @param out    成品图片输出流
     * @param x      开始位置的x坐标
     * @param y      开始位置的y坐标
     * @param width  裁剪的宽度
     * @param height 裁剪的高度
     * @throws IOException
     */
    public static void cut(InputStream src, OutputStream out, int x, int y, int width, int height) throws IOException {
        Iterator<ImageReader> iterator = ImageIO.getImageReadersByFormatName("png");
        ImageReader reader = iterator.next();
        ImageInputStream iis = ImageIO.createImageInputStream(src);
        reader.setInput(iis, true);
        ImageReadParam param = reader.getDefaultReadParam();
        Rectangle rect = new Rectangle(x, y, width, height);
        param.setSourceRegion(rect);
        BufferedImage bi = reader.read(0, param);
        ImageIO.write(bi, "png", out);
    }
}
