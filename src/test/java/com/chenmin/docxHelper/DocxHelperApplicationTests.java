package com.chenmin.docxHelper;

import com.chenmin.docxHelper.service.DocxGenerationService;
import com.chenmin.docxHelper.service.DocxMergingNewService;
import com.chenmin.docxHelper.service.DocxMergingService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;

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

	@Test
	public void testService() throws Exception {
		docxMergingNewService.merge();
	}
}
