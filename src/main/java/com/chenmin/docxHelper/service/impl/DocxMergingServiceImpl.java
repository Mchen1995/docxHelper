package com.chenmin.docxHelper.service.impl;

import com.chenmin.docxHelper.service.DocxMergingService;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service("docxMergingService")
public class DocxMergingServiceImpl implements DocxMergingService {

    /**
     * LINUX 上的文档路径
     */
    public static final String DIR_LINUX = "/cib/docxTools/testReport/";

    /**
     * MACOS 上的文档路径
     */
    public static final String DIR_MAC_OS = "/Users/chenmin/Desktop/docxHelper/";

    @Override
    public void mergeDocx() {
        String version = "2.4.5.6";
        List<File> fileList = new ArrayList<>();
        String targetFilename = "集团门户技术测试报告_" + version + ".docx";
        String firstFilename = "集团门户技术测试报告.docx";
        File directory = new File(DIR_MAC_OS);
        File[] files = directory.listFiles();
        File targetFile = new File(DIR_MAC_OS + targetFilename);  //
        assert files != null;
        fileList.add(new File(DIR_MAC_OS + firstFilename));
        for (File file : files) {
            String name = file.getName();
            if (name.endsWith(".docx") && !name.equals(firstFilename)) {
                System.out.println("识别到文件：" + name);
                fileList.add(file);
            }
        }
        appendDocx(targetFile, fileList);
        System.out.println("生成文件：" + targetFilename);
    }

    private void appendDocx(File outfile, List<File> targetFile) {
        try {
            OutputStream dest = new FileOutputStream(outfile);
            ArrayList<XWPFDocument> documentList = new ArrayList<>();
            XWPFDocument doc = null;
            for (File file : targetFile) {
                FileInputStream in = new FileInputStream(file.getPath());
                OPCPackage open = OPCPackage.open(in);
                XWPFDocument document = new XWPFDocument(open);
                documentList.add(document);
            }
            for (int i = 0; i < documentList.size(); i++) {
                doc = documentList.get(0);
                if (i != 0) {
                    //解决word合并完后，所有表格都紧紧挨在一起，没有分页。加上了分页符可解决
                    documentList.get(i).createParagraph().setPageBreak(true);
                    appendBody(doc, documentList.get(i));
                }
            }
            assert doc != null;
            doc.write(dest);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void appendBody (XWPFDocument src, XWPFDocument append) throws Exception {
        CTBody src1Body = src.getDocument().getBody();
        CTBody src2Body = append.getDocument().getBody();

        List<XWPFPictureData> allPictures = append.getAllPictures();
        // 记录图片合并前及合并后的ID
        Map<String, String> map = new HashMap<>();
        for (XWPFPictureData picture : allPictures) {
            String before = append.getRelationId(picture);
            //将原文档中的图片加入到目标文档中
            String after = src.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
            map.put(before, after);
        }
        appendBody(src1Body, src2Body, map);
    }

    private static void appendBody(CTBody src, CTBody append, Map<String, String> map) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);

        String srcString = src.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String surfix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        //下面这部分可以去掉，我加上的原因是合并的时候，有时候出现打不开的情况，对照document.xml将某些标签去掉就可以正常打开了
        addPart = addPart.replaceAll("w14:paraId=\"[A-Za-z0-9]{1,10}\"", "");
        addPart = addPart.replaceAll("w14:textId=\"[A-Za-z0-9]{1,10}\"", "");
        addPart = addPart.replaceAll("w:rsidP=\"[A-Za-z0-9]{1,10}\"", "");
        addPart = addPart.replaceAll("w:rsidRPr=\"[A-Za-z0-9]{1,10}\"", "");
        addPart = addPart.replace("<w:headerReference r:id=\"rId8\" w:type=\"default\"/>", "");
        addPart = addPart.replace("<w:footerReference r:id=\"rId9\" w:type=\"default\"/>", "");
        addPart = addPart.replace("xsi:nil=\"true\"", "");

        if (map != null && !map.isEmpty()) {
            //对xml字符串中图片ID进行替换
            for (Map.Entry<String, String> set : map.entrySet()) {
                addPart = addPart.replace(set.getKey(), set.getValue());
            }
        }
        //将两个文档的xml内容进行拼接
        CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart + surfix);

        src.set(makeBody);
    }
}
