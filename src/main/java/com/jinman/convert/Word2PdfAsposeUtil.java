package com.jinman.convert;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;

public class Word2PdfAsposeUtil {


    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = Word2PdfAsposeUtil.class.getClassLoader().getResourceAsStream("license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    public static boolean doc2pdf(String inPath, String outPath) {
        if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档会有水印产生
            return false;
        }
        FileOutputStream os = null;
        try {
            long old = System.currentTimeMillis();
            File file = new File(outPath); // 新建一个空白pdf文档
            os = new FileOutputStream(file);
            Document doc = new Document(inPath); // Address是将要被转化的word文档
            doc.save(os, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF,
            // EPUB, XPS, SWF 相互转换
            long now = System.currentTimeMillis();
            System.out.println("pdf转换成功，共耗时：" + ((now - old) / 1000.0) + "秒"); // 转化用时
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }finally {
            if (os != null) {
                try {
                    os.flush();
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }

    public static void main(String[] arg){
        System.out.println("被转化的word文档地址：");
        Scanner scanner = new Scanner(System.in);
        String docPath = scanner.next();
        System.out.println("pdf被输出的地址：");
        String pdfPath = scanner.next();
//        String docPath = "C:\\Users\\fengjinman\\Desktop\\接口文档(1).docx";
//        String pdfPath = "C:\\Users\\fengjinman\\Desktop\\接口文档(1).pdf";
//        String docPath = "C:\\Users\\fengjinman\\Desktop\\03.统一消息中心数据字典v1.0(1).doc";
//        String pdfPath = "C:\\Users\\fengjinman\\Desktop\\03.统一消息中心数据字典v1.0(1).pdf";
        Word2PdfAsposeUtil.doc2pdf(docPath,pdfPath);
    }

}
