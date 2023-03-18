package com.chenmin.docxHelper.service;

import java.io.IOException;

public interface DocxGenerationService {
    /**
     * 生成技术测试报告的汇总表格
     */
    void generateTestReportSummary() throws IOException;

    /**
     * 生成技术测试报告的案例表格
     */
    void generateTestReportCase() throws IOException;
}
