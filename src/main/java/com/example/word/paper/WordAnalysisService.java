package com.example.word.paper;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.util.List;

/**
 * 试卷解析流程
 * @author liujunhui
 * @date 2021/2/7 17:56
 */
public class WordAnalysisService {

    public static void main(String[] args) {
        String path = "";

    }

    /**
     *
     * @param wordFile
     * @return
     */
    public boolean analysisWordPaper(String wordFile) {
        return true;
    }

    /**
     * 筛选目录下所有word文件
     * @param parentPath
     * @return
     */
    public List<File> listWordFiles(String parentPath) {
        File parentFile = new File(parentPath);
        if (!parentFile.isDirectory()) {
            return null;
        }
        List<File> fileList = (List<File>) FileUtils.listFiles(new File(parentPath), null, true);
        return null;
    }
}
