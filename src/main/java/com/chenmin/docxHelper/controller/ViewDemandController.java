package com.chenmin.docxHelper.controller;

import com.chenmin.docxHelper.model.DemandVO;
import com.chenmin.docxHelper.service.DocxGenerationService;
import com.chenmin.docxHelper.service.UIService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.util.List;


@Controller
public class ViewDemandController {

    @Autowired
    private UIService uiService;

    @Autowired
    private DocxGenerationService docxGenerationService;


    @GetMapping("/view")
    public String getDemands(Model model) throws IOException {
        return uiService.viewDemands(model);
    }

    @PostMapping("/gen")
    @ResponseBody
    public String updateScores(@RequestBody List<DemandVO> demandList) throws IOException {
        System.out.println(demandList.size());
        System.out.println(demandList.get(0).getCaseCount());
        System.out.println(demandList.get(0).getDescription());
//        docxGenerationService.generateTestReportCase();
        return "success";
    }

}
