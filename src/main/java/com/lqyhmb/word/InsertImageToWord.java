package com.lqyhmb.word;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * 通过 poi 往 word插入图片
 * Created by Rodriguez
 * 2018/3/7 10:08
 * url: https://stackoverflow.com/questions/26764889/how-to-insert-a-image-in-word-document-with-apache-poi
 * url: http://blog.csdn.net/ro_wsy/article/details/24673333
 */
public class InsertImageToWord {


    //网络图片
    @Test
    public void test2() throws IOException, InvalidFormatException {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph title = doc.createParagraph(); //
        XWPFRun run = title.createRun(); // XWPFRun对象用一组公共属性定义文本区域。
        run.setText("Test Web Image");
        run.setBold(true); // 字体变粗
        title.setAlignment(ParagraphAlignment.CENTER); // 字体居中

        // 获取网络图片输入流
        URL url = new URL("http://img.taopic.com/uploads/allimg/120727/201995-120HG1030762.jpg");
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5 * 1000);
        InputStream inStream = conn.getInputStream();//通过输入流获取图片数据

        // 插入图片
        run.addPicture(inStream, XWPFDocument.PICTURE_TYPE_JPEG, "1", Units.toEMU(200), Units.toEMU(200)); // 200*200 pixels
        run.addBreak();
        run.setText("web image");
        FileOutputStream fos = new FileOutputStream("H:\\webImage.docx");
        doc.write(fos);
        fos.close();
    }


    // 本地图片
    @Test
    public void test() throws IOException, InvalidFormatException {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph title = doc.createParagraph(); //
        XWPFRun run = title.createRun(); // XWPFRun对象用一组公共属性定义文本区域。
        run.setText("Fig.1 A Natural Scene");
        run.setBold(true); // 字体变粗
        title.setAlignment(ParagraphAlignment.CENTER); // 字体居中

        String imgFile = "H:\\图片1.jpg";
        String imgFile2 = "H:\\图片2.jpg";
        FileInputStream is = new FileInputStream(imgFile);
        FileInputStream is2 = new FileInputStream(imgFile2);
        run.addBreak(); // 换行

        // 获取网络图片输入流
        /*URL url = new URL("http://img.taopic.com/uploads/allimg/120727/201995-120HG1030762.jpg");
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        conn.setRequestMethod("GET");
        conn.setConnectTimeout(5 * 1000);
        InputStream inStream = conn.getInputStream();//通过输入流获取图片数据*/

        run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, "1", Units.toEMU(200), Units.toEMU(200)); // 200*200 pixels
        run.addPicture(is2, XWPFDocument.PICTURE_TYPE_JPEG, "2", Units.toEMU(200), Units.toEMU(200)); // 200*200 pixels

        //run.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU(220), Units.toEMU(150)); // 200*200 pixels
        //run.addPicture(is2, XWPFDocument.PICTURE_TYPE_JPEG, imgFile2, Units.toEMU(220), Units.toEMU(150)); // 200*200 pixels
        run.addBreak();
        run.setText("东进口车道分配情况");
        run.addTab(); // 空格
        run.addTab(); // 空格
        run.addTab(); // 空格
        run.setText("东进口灯具使用情况");

        FileOutputStream fos = new FileOutputStream("H:\\test.docx");
        doc.write(fos);
        fos.close();
    }

}
