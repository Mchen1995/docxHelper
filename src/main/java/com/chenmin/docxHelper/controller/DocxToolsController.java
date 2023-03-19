package com.chenmin.docxHelper.controller;

import com.chenmin.docxHelper.model.DemandVO;
import com.chenmin.docxHelper.service.*;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@Controller
public class DocxToolsController {

    @Autowired
    private UIService uiService;

    @Autowired
    private DocxGenerationService docxGenerationService;

    @Autowired
    private ExcelHelperService excelHelperService;

    @Autowired
    private DocxMergingService docxMergingService;

    @Autowired
    private DocxMergingNewService docxMergingNewService;

    @GetMapping("/merge")
    @ResponseBody
    public String mergeDocx() throws Exception {
//        docxMergingService.mergeDocx();
        docxMergingNewService.merge();
        return "success";
    }


    @GetMapping("/list")
    public String getDemands(Model model) throws IOException {
        return uiService.viewDemands(model);
    }

    @GetMapping("/gen/summary/docx")
    @ResponseBody
    public String genSummaryDocx() throws IOException {
        docxGenerationService.generateTestReportSummary();
        return "success";
    }

    @PostMapping("/gen/case/docx")
    @ResponseBody
    public String genCaseDocx(@RequestBody List<DemandVO> demandList) throws IOException {
        String name = demandList.get(0).getName();
        List<List<String>> lists = excelHelperService.readExcel();
        List<String> idList = new ArrayList<>();
        List<String> caseCountList = new ArrayList<>();
        List<String> orderList = new ArrayList<>();

        for (int i = 0; i < demandList.size(); i++) {
            int caseCount = demandList.get(i).getCaseCount();
            if (caseCount >0) {
                idList.add(lists.get(1).get(i));
                caseCountList.add(String.valueOf(caseCount));
                orderList.add(lists.get(0).get(i));
            }
        }
        docxGenerationService.generateTestReportCase(idList, caseCountList, orderList, name);
        return "success";
    }

    @GetMapping("/list/{name}")
    public void downloadFile(@PathVariable String name, HttpServletResponse response) throws IOException, InvalidFormatException {
        // 从文件系统中获取文件输入流
        File file = new File(name + "_技术测试报告.docx");
        String filename = file.getName();
        FileInputStream fis = new FileInputStream(file);
        response.setHeader("Content-Disposition", "attachment; filename=" +
                new String(filename.getBytes(), "ISO-8859-1"));
        response.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
        document.write(response.getOutputStream());
    }

    @GetMapping("/merge/download")
    public void downloadMergeFile(HttpServletResponse response) throws IOException, InvalidFormatException {
        // 从文件系统中获取文件输入流
        File file = new File("集团门户技术测试报告_版本号.docx");
        String filename = file.getName();
        FileInputStream fis = new FileInputStream(file);
        response.setHeader("Content-Disposition", "attachment; filename=" +
                new String(filename.getBytes(), "ISO-8859-1"));
        response.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
        document.write(response.getOutputStream());
    }

}
