package com.chenmin.docxHelper.service.impl;

import com.chenmin.docxHelper.service.DocxMergingNewService;
import org.apache.commons.lang3.StringUtils;
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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.chenmin.docxHelper.service.impl.DocxMergingServiceImpl.DIR_MAC_OS;

@Service("docxMergingNewService")
public class DocxMergingNewServiceImpl implements DocxMergingNewService {
    @Override
    public void merge(String filename1, String filename2) throws Exception {
        InputStream in1 = null;
        InputStream in2 = null;
        String version = "版本号";
        File dest = new File("集团门户技术测试报告_" + version + ".docx");
        try {
            in1 = new FileInputStream(filename1);
            in2 = new FileInputStream(filename2);
        } catch (Exception e) {
            e.printStackTrace();
        }
        mergeWord(in1, in2, dest);
    }

    @Override
    public void merge() throws Exception {
        File directory = new File(DIR_MAC_OS);
        File[] files = directory.listFiles();
        assert files != null;
        List<File> allFiles = new ArrayList<>();
        for (File file : files) {
            String name = file.getName();
            if (name.endsWith(".docx") && name.contains("_技术测试报告")) {
                System.out.println("识别到文件：" + name);
                allFiles.add(file);
            }
        }

        List<String> filenames = new ArrayList<>();
        for (int i=1; i< allFiles.size();i++) {
            filenames.add(allFiles.get(i).getName());
        }

        InputStream firstInputStream = new FileInputStream(allFiles.get(0).getName());
        String version = "版本号";
        File dest = new File("集团门户技术测试报告_" + version + ".docx");
        List<InputStream> inputStreamList = new ArrayList<>();
        for (String filename : filenames) {
            InputStream in = new FileInputStream(filename);
            inputStreamList.add(in);
        }
        mergeWordList(firstInputStream, inputStreamList, dest);
    }

    private static void mergeWordList(InputStream firstInputStream, List<InputStream> inputStreamList, File outputFile) throws Exception {
        OPCPackage src1Package = OPCPackage.open(firstInputStream);
        OutputStream dest = new FileOutputStream(outputFile);

        List<XWPFDocument> documentList = new ArrayList<>();
        XWPFDocument src1Document = new XWPFDocument(src1Package);
        for (InputStream inputStream : inputStreamList) {
            OPCPackage opcPackage = OPCPackage.open(inputStream);
            XWPFDocument document = new XWPFDocument(opcPackage);
            documentList.add(document);
        }

        for (XWPFDocument document : documentList) {
            appendBody(src1Document, document);
        }
        src1Document.write(dest);
    }

    private static void mergeWord(InputStream firstInputStream, InputStream secondInputStream, File outputFile) throws Exception {
        OPCPackage src1Package = null;
        OPCPackage src2Package = null;
        OutputStream dest = new FileOutputStream(outputFile);
        try {
            src1Package = OPCPackage.open(firstInputStream);
            src2Package = OPCPackage.open(secondInputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
        assert src1Package != null;
        XWPFDocument src1Document = new XWPFDocument(src1Package);
        assert src2Package != null;
        XWPFDocument src2Document = new XWPFDocument(src2Package);
        appendBody(src1Document, src2Document);
        src1Document.write(dest);
    }

    public static void appendBody(XWPFDocument src, XWPFDocument append) throws Exception {
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
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));

        if (map != null && !map.isEmpty()) {
            //对xml字符串中图片ID进行替换 z
            // 下面注释掉的方式会发生图片id冲突
            // for (Map.Entry<String, String> set : map.entrySet()) {
            //     addPart = addPart.replace(set.getKey(), "RE:"+set.getValue());
            // }
            // addPart = addPart.replaceAll("RE:",");

            // 采用正则追加替换方式完美解决
            String patter = StringUtils.join(map.keySet(), "|");
            Pattern compile = Pattern.compile(patter);
            Matcher matcher = compile.matcher(addPart);
            StringBuilder sb = new StringBuilder();
            while (matcher.find()) {
                String value = map.get(matcher.group());
                if (value != null) {
                    matcher.appendReplacement(sb, value);
                }
            }
            matcher.appendTail(sb);
            addPart = sb.toString();
        }
        //将两个文档的xml内容进行拼接
        CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
        src.set(makeBody);
    }
}
