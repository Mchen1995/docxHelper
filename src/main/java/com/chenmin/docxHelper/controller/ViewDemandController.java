package com.chenmin.docxHelper.controller;

import com.chenmin.docxHelper.service.UIService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.io.IOException;


@Controller
public class ViewDemandController {

    @Autowired
    private UIService uiService;


    @GetMapping("/view")
    public String getDemands(Model model) throws IOException {
        return uiService.viewDemands(model);
    }
}
