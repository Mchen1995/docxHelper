package com.chenmin.docxHelper.service.impl;

import com.chenmin.docxHelper.service.ExcelHelperService;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service("excelHelperService")
public class ExcelHelperServiceImpl implements ExcelHelperService {

    public static final String EXCEL_FILE_IN_MAC_OS = "软件下发需求.xls";

    /**
     * Windows 上的需求清单文件路径
     */
    public static final String EXCEL_FILE_IN_WIN = "D:\\genDocx\\软件下发需求.xls";

    @Override
    public List<List<String>> readExcel() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(EXCEL_FILE_IN_MAC_OS);
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        List<String> orderList = new ArrayList<>();  // 序号
        List<String> idList = new ArrayList<>();  // 需求编号
        List<String> descriptionList = new ArrayList<>();  // 需求描述
        List<List<String>> res = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            orderList.add(String.valueOf(i));
            idList.add(row.getCell(1).toString());
            descriptionList.add(row.getCell(2).toString());
        }
        res.add(orderList);
        res.add(idList);
        res.add(descriptionList);
        fileInputStream.close();
        return res;
    }
}
