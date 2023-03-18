package com.chenmin.docxHelper.service;

import org.springframework.ui.Model;

import java.io.IOException;

public interface UIService {
    /**
     * 查看 Excel
     * @param model 模型
     * @return 返回 HTML 文件名
     */
    String viewDemands(Model model) throws IOException;
}
