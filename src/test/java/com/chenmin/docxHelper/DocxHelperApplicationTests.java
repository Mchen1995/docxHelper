package com.chenmin.docxHelper;

import com.chenmin.docxHelper.service.DocxGenerationService;
import com.chenmin.docxHelper.service.DocxMergingService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.IOException;

@SpringBootTest
class DocxHelperApplicationTests {

	@Autowired
	private DocxGenerationService docxGenerationService;

	@Autowired
	private DocxMergingService docxMergingService;

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
}
