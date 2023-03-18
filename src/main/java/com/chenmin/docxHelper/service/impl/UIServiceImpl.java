package com.chenmin.docxHelper.service.impl;

import com.chenmin.docxHelper.model.DemandVO;
import com.chenmin.docxHelper.service.ExcelHelperService;
import com.chenmin.docxHelper.service.UIService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.ui.Model;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service("UIService")
public class UIServiceImpl implements UIService {

    public static final String HTML_FILE_NAME = "demandList";

    @Autowired
    private ExcelHelperService excelHelperService;

    @Override
    public String viewDemands(Model model) throws IOException {
        List<List<String>> dataSrc = excelHelperService.readExcel();
        List<String> orderList = dataSrc.get(0);
        int rows = orderList.size();
        List<String> idList = dataSrc.get(1);
        List<String> descriptionList = dataSrc.get(2);

        List<DemandVO> demandList = new ArrayList<>();
        for (int i = 0; i < rows; i++) {
            demandList.add(new DemandVO(
                    Integer.parseInt(orderList.get(i)),
                    idList.get(i),
                    descriptionList.get(i),
                    0));
        }
        model.addAttribute(HTML_FILE_NAME, demandList);
        return HTML_FILE_NAME;
    }
}
