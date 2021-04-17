package com.example.word.aspose;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;

import javax.swing.filechooser.FileSystemView;
import java.io.File;

/**
 * @author liujunhui
 * @date 2021/1/25 11:46
 */
public class HtmlToWord {
    private static String outPath = FileSystemView.getFileSystemView().getHomeDirectory().getPath();

    public static void main(String[] args) throws Exception {
        Document document = new Document();
        DocumentBuilder db = new DocumentBuilder(document);
        db.insertHtml("（3分）一幅如图甲所示的漫画立在桌面上。小霞把一个装水的玻璃杯放在漫画前，惊奇地发现：\" +\n" +
                "                \"透过水杯看到漫画中的老鼠变“胖”了，还掉头奔向猫，如图乙。小霞观察分析：装水的圆柱形玻璃杯横切面中间厚\" +\n" +
                "                \"，边缘薄，起到（填“平面镜”、“凸透镜”或“凹透镜” \" +\n" +
                "                \"$ ) $\" +\n" +
                "                \"的作用，使图片横向放大、颠倒，透过水杯看到的是老鼠的（填“实”或“虚” \" +\n" +
                "                \"$ ) $\" +\n" +
                "                \"像，此时老鼠到玻璃杯的距离满足（填“\" +\n" +
                "                \"$ f&lt;u&lt;2f $”或“$ u&gt;2f $” $ ) $。<br/> ");
        document.save(outPath + File.separator + "123.docx");
    }
}
