package com.chenmin.docxHelper;

import com.chenmin.docxHelper.service.DocxGenerationService;
import com.chenmin.docxHelper.service.DocxMergingNewService;
import com.chenmin.docxHelper.service.DocxMergingService;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlOptions;
import org.apache.xmlbeans.impl.xb.xsdschema.PatternDocument;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SpringBootTest
class DocxHelperApplicationTests {

	@Autowired
	private DocxGenerationService docxGenerationService;

	@Autowired
	private DocxMergingService docxMergingService;

	@Autowired
	private DocxMergingNewService docxMergingNewService;

	@Test
	void contextLoads() {
	}

	/**
	 * 生成技术测试报告汇总表格测试
	 * @throws IOException
	 */
	@Test
	void genSummaryDocx() throws IOException {
		docxGenerationService.generateTestReportSummary();
		System.out.println(1);
	}

	/**
	 * 生成技术测试报告案例表格测试
	 * @throws IOException
	 */
	@Test
	void genCaseDocx() throws IOException {
//		docxGenerationService.generateTestReportCase();
		System.out.println(1);
	}

	/**
	 * 合并测试
	 * @throws IOException
	 */
	@Test
	void mergeDocx() throws IOException {
		docxMergingService.mergeDocx();
		System.out.println(1);
	}

	@Test
	void mergeDocxNew2() throws Exception {
		InputStream in1 = null;
		InputStream in2 = null;
		File dest = new File("merge.docx");
		try {
			in1 = new FileInputStream("集团门户技术测试报告.docx");
			in2 = new FileInputStream("陈敏_技术测试报告.docx");
		} catch (Exception e) {
			e.printStackTrace();
		}
		mergeWord(in1, in2, dest);
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

	@Test
	void mergeDocxNew3() throws IOException {
		XWPFTemplate template = XWPFTemplate.compile("集团门户技术测试报告.docx").render(
				new HashMap<String, Object>(){{
					put("title", "Hi, poi-tl Word模板引擎");
					put("docx_word1", new DocxRenderData(new File("陈奕迅_技术测试报告.docx")));
					put("docx_word2", new DocxRenderData(new File("陈敏_技术测试报告.docx")));
				}});

		FileOutputStream out = new FileOutputStream("merge.docx");
		template.write(out);
		out.flush();
		out.close();
		template.close();
	}

	@Test
	public void testService() throws Exception {
		docxMergingNewService.merge();
	}
}
