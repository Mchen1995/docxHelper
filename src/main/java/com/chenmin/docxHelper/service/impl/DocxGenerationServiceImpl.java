package com.chenmin.docxHelper.service.impl;

import com.chenmin.docxHelper.service.DocxGenerationService;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service("docxGenerationService")
public class DocxGenerationServiceImpl implements DocxGenerationService {

    /**
     * 需求勾选标志
     */
    public static final String PICK_FLAG = "a";

    /**
     * Mac 上的需求清单文件路径
     */
    public static final String EXCEL_FILE_IN_MAC_OS = "/Users/chenmin/Desktop/docxHelper/src/main/resources/软件下发需求.xls";

    /**
     * Windows 上的需求清单文件路径
     */
    public static final String EXCEL_FILE_IN_WIN = "D:\\genDocx\\软件下发需求.xls";

    /**
     * 表格宽度
     */
    public static final Integer TABLE_WIDTH = 8000;

    /**
     * 单元格宽度1500
     */
    public static final String CELL_WIDTH_1500 = "1500";

    /**
     * 单元格宽度1000
     */
    public static final String CELL_WIDTH_1000 = "1000";

    /**
     * 单元格宽度500
     */
    public static final String CELL_WIDTH_500 = "500";

    /**
     * 一般行高度
     */
    public static final Integer COMMON_HEIGHT = 400;
    /**
     * 行高度（截图）
     */
    public static final Integer PICTURE_HEIGHT = 4000;

    @Override
    public void generateTestReportSummary() throws IOException {
        String targetFilename = "集团门户技术测试报告.docx";
        XWPFDocument document = new XWPFDocument();
        addHeader(document);
        List<String> ids = readExcelForSummarizedTable();
        createSummarizedTable(document, ids);
        FileOutputStream out = new FileOutputStream(targetFilename);
        document.write(out);
        out.close();
        System.out.println("生成文档：" + targetFilename);
    }

    @Override
    public void generateTestReportCase(List<String> numberList,
                                       List<String> caseCountList,
                                       List<String> orderList,
                                       String name) throws IOException {
        XWPFDocument document = new XWPFDocument();  // 创建全文
        createCaseTable(document, numberList, caseCountList, orderList, name);
        String targetFilename = name + "_技术测试报告.docx";
        FileOutputStream out = new FileOutputStream(targetFilename);
        document.write(out);
        out.close();
        System.out.println("生成文档：" + targetFilename);
    }

    /**
     * 生成案例表格
     * @param document 文档
     * @param ids 需求编号
     * @param caseCountList 案例个数
     * @param orderList 序号
     * @param name 姓名
     */
    private void createCaseTable(XWPFDocument document,
                                       List<String> ids,
                                       List<String> caseCountList,
                                       List<String> orderList,
                                       String name) {
        for (int i = 0; i < ids.size(); i++) {
            int numberOfCase = Integer.parseInt(caseCountList.get(i));
            int order = Integer.parseInt(orderList.get(i));
            for (int j = 0; j < numberOfCase; j++) {
                XWPFTable table = document.createTable(9, 4);
                // 水平合并单元格
                mergeHorizontal(table, 2, 1, 3);
                mergeHorizontal(table, 3, 1, 3);
                mergeHorizontal(table, 4, 0, 3);
                mergeHorizontal(table, 5, 0, 3);
                mergeHorizontal(table, 6, 0, 3);
                mergeHorizontal(table, 7, 1, 3);
                mergeHorizontal(table, 8, 1, 3);
                table.setWidth(TABLE_WIDTH);
                setCaseTableRow1(table, name);
                setCaseTableRow2(table, order + "-" + (j + 1));
                setCaseTableRowOthers(table);
                addBreak(document);
            }
        }
    }

    private void setCaseTableRow1(XWPFTable table, String name) {
        XWPFTableRow row0 = table.getRow(0);
        row0.getCell(0).setWidth(CELL_WIDTH_1000);
        row0.setHeight(COMMON_HEIGHT);
        row0.getCell(0).setText("测试人");
        row0.getCell(1).setText(name);
        row0.getCell(2).setText("编写人");
        row0.getCell(3).setText(name);
    }

    /**
     * 赋值测试案例表格第二行
     * @param table      表格
     * @param caseNumber 用例编号
     */
    private void setCaseTableRow2(XWPFTable table, String caseNumber) {
        XWPFTableRow row1 = table.getRow(1);
        row1.setHeight(COMMON_HEIGHT);
        row1.getCell(0).setText("用例编号");
        row1.getCell(1).setText(caseNumber);
        row1.getCell(2).setText("测试日期");
        row1.getCell(3).setText("YYYY-MM-DD");
    }

    private void setCaseTableRowOthers(XWPFTable table) {
        XWPFTableRow row2 = table.getRow(2);
        row2.getCell(0).setText("用例名称");
        row2.setHeight(COMMON_HEIGHT);

        table.getRow(3).getCell(0).setText("测试目标");
        table.getRow(3).setHeight(COMMON_HEIGHT);

        table.getRow(4).getCell(0).setText("输入说明：");
        table.getRow(4).setHeight(COMMON_HEIGHT);

        table.getRow(5).getCell(0).setText("预期结果：测试通过");
        table.getRow(5).setHeight(COMMON_HEIGHT);

        table.getRow(6).getCell(0).setText("实际结果（截图）：");
        table.getRow(6).setHeight(PICTURE_HEIGHT);

        table.getRow(7).getCell(0).setText("结果及意见");
        table.getRow(7).getCell(1).setText("测试通过");
        table.getRow(7).setHeight(COMMON_HEIGHT);

        table.getRow(8).getCell(0).setText("备注");
        table.getRow(8).setHeight(COMMON_HEIGHT);
    }

    private void mergeHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    private List<List<String>> readExcelForCaseTable() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(EXCEL_FILE_IN_MAC_OS);
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        List<String> ids = new ArrayList<>();
        List<String> orderList = new ArrayList<>();
        List<List<String>> res = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            orderList.add(String.valueOf(i));  // 序号
            ids.add(row.getCell(1).toString());  // 需求编号列表
        }
        res.add(orderList);
        res.add(ids);
        fileInputStream.close();
        return res;
    }

    /**
     * 技术测试报告的文档前序部分
     * @param document 文档
     */
    private void addHeader(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setText("集团网络金融门户系统-【下发编号】");
        run.setBold(true);
        run.setFontSize(20);

        XWPFParagraph paragraph2 = document.createParagraph();
        paragraph2.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run2 = paragraph2.createRun();
        run2.setText("技术测试报告");
        run2.setBold(true);
        run2.setFontSize(20);

        document.createParagraph().createRun().setText("一、测试日期：YYYY年MM月DD日-MM月DD日");
        document.createParagraph().createRun().setText("二、测试人员：苏培泳、曹喜阳、黄梅兰、高发鑫、谢淑慧、张湛、宋浦榕、郑诗渊、吴铁浩、黄文龙、陈敏");
        document.createParagraph().createRun().setText("三、测试结果：测试通过");
    }

    /**
     * 读取需求清单的需求编号
     * @return 返回列表
     */
    private List<String> readExcelForSummarizedTable() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(EXCEL_FILE_IN_MAC_OS);
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet sheet = workbook.getSheetAt(0);
        List<String> ids = new ArrayList<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            HSSFRow row = sheet.getRow(i);
            ids.add(row.getCell(1).toString());
        }
        return ids;
    }

    /**
     * 创建技术测试报告的汇总表格
     * @param document 文档
     * @param ids 需求编号列表
     */
    private void createSummarizedTable(XWPFDocument document, List<String> ids) {
        int size = ids.size();
        XWPFTable table = document.createTable(size + 1, 5);
        table.setWidth(TABLE_WIDTH);
        XWPFTableRow row = table.getRow(0);
        row.setHeight(COMMON_HEIGHT);
        row.getCell(0).setText("序号");
        row.getCell(0).setWidth(CELL_WIDTH_500);

        row.getCell(1).setText("需求编号");
        row.getCell(1).setWidth(CELL_WIDTH_1500);

        row.getCell(2).setText("用例数");
        row.getCell(2).setWidth(CELL_WIDTH_500);

        row.getCell(3).setText("用例号");
        row.getCell(3).setWidth(CELL_WIDTH_500);

        row.getCell(4).setText("测试结果");
        row.getCell(4).setWidth(CELL_WIDTH_500);
        for (int i = 0; i < size; i++) {
            table.getRow(i+1).getCell(0).setText(String.valueOf(i+1));  // 序号
            table.getRow(i+1).getCell(1).setText(ids.get(i));  // 需求编号
            table.getRow(i+1).getCell(4).setText("通过");  // 测试结果
        }
        addBreak(document);
    }

    /**
     * 新增 2 个空行
     * @param document 文档
     */
    private void addBreak(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("\n\n");
    }
}
