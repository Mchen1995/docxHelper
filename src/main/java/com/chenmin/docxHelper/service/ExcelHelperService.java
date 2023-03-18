package com.chenmin.docxHelper.service;

import java.io.IOException;
import java.util.List;

public interface ExcelHelperService {
    /**
     * 读取需求清单 Excel
     * @return 返回读取的数据
     */
    List<List<String>> readExcel() throws IOException;
}
