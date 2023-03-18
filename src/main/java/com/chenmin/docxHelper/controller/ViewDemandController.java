package com.chenmin.docxHelper.controller;

import com.chenmin.docxHelper.model.DemandVO;
import com.chenmin.docxHelper.service.DocxGenerationService;
import com.chenmin.docxHelper.service.ExcelHelperService;
import com.chenmin.docxHelper.service.UIService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


@Controller
public class ViewDemandController {

    @Autowired
    private UIService uiService;

    @Autowired
    private DocxGenerationService docxGenerationService;

    @Autowired
    private ExcelHelperService excelHelperService;


    @GetMapping("/view")
    public String getDemands(Model model) throws IOException {
        return uiService.viewDemands(model);
    }

    @PostMapping("/gen/case/docx")
    @ResponseBody
    public String genCaseDocx(@RequestBody List<DemandVO> demandList) throws IOException {
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
        docxGenerationService.generateTestReportCase(idList, caseCountList, orderList);
        return "success";
    }

}
