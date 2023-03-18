package com.chenmin.docxHelper.controller;

import com.chenmin.docxHelper.service.DocxGenerationService;
import com.chenmin.docxHelper.service.DocxMergingService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("api/docx")
public class DocxHelperController {
    @Autowired
    private DocxGenerationService docxGenerationService;

    @Autowired
    private DocxMergingService docxMergingService;

    @GetMapping("/merge")
    public void mergeDocx() {
        docxMergingService.mergeDocx();
    }

    @GetMapping("/gen/test/summary")
    public void generateTestReportSummary() throws IOException {
        docxGenerationService.generateTestReportSummary();
    }
}
