package com.jac.service;

import com.jac.model.DataCell;
import com.jac.model.DataCellMergeInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author ：jacketzc
 * Created in 2020/11/12 9:54
 */
public class ReportModelService {

    private CellStyle cellStyle;

    /**
     * 百世报表导出（导出所有的sheet）
     * @param allData 全部的数据需要封装为 map格式，map的 key为 sheet的名称，value的格式应该为 List<List>，对应着单个sheet的数据
     * @param file 需要输出的目录
     * @throws Exception
     */
    public void exportExcel(Map<String, List> allData, File file) throws Exception {
        ExcelUtils excelUtils = new ExcelUtils();

        //处理数据，将原生的数据转换为包含合并信息的 DataCell
        allData.forEach((k,v)->{
            allData.put(k, excelUtils.addMergeInfo(k, v));
        });


        Workbook workbook = excelUtils.createEmptyXlsxWorkbook();
        //初始化样式
        this.cellStyle = createCellStyle(workbook);

        //设置每一个sheet，此处操作的目的不在于拆分业务与逻辑，只是代码太长了，拆一下

        createSheet1(workbook, allData.get("集团汇总"));
        createSheet2(workbook, allData.get("分公司汇总"));
        createSheet3(workbook, allData.get("快递事业部"));
        createSheet4(workbook, allData.get("快运事业部"));
        createSheet5(workbook, allData.get("供应链事业部"));
        createSheet6(workbook, allData.get("店加事业部"));

        //输出到文件
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }

    /**
     * 创建样式
     * @param workbook
     * @return
     */
    private CellStyle createCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        //居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //自动换行
        cellStyle.setWrapText(true);
        //字体
        Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 10);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * “集团汇总”的sheet
     * 因为表格太大了，故将生成每个sheet的方法拆分
     */
    private void createSheet1(Workbook workbook, List<List> data) {

        //创建固定的表头
        Sheet sheet1 = workbook.createSheet("集团汇总");
        AtomicInteger row1num = new AtomicInteger();
        //第一行
        Row row0 = sheet1.createRow(row1num.getAndIncrement());
        Cell title = row0.createCell(0);
        title.setCellValue("百世集团EHS双周报");
        sheet1.addMergedRegion(new CellRangeAddress(0, 0, 0, 54));
        title.setCellStyle(cellStyle);
        //第二行
        Row row1 = sheet1.createRow(row1num.getAndIncrement());
        String[] row_first = {"地区", "集团/事业部", "在职人数", "总部办公人数", "一线人数", "未遂事件数量",
                "考核工伤事故 月度目标0.23‰", "", "", "", "消防事故 月度目标 0", "",
                "财产损失安全事故 月度目标 0", "", "", "环境污染 0", "", "烧车事故", "", "食品安全事故 0（店加）",
                "", "政府检查罚款", "", "", "安全隐患", "", "", "", "", "", "", "",
                "特种设备", "", "", "", "特种作业", "", "", "", "新改扩建场所安全评估",
                "安全培训 新员工安全培训覆盖率（入职7日内） 100%", "", "", "", "", "", "", "", "", "",
                "安全人员数量", "", "工业风扇", "备注"};
        for (int i = 0; i < row_first.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(row_first[i]);
            cell.setCellStyle(cellStyle);
        }
        //合并1-6
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 3, 3));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 4, 4));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 5, 5));
        //合并7-24
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 6, 9));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 10, 11));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 12, 14));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 15, 16));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 17, 18));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 19, 20));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 21, 23));
        //合并25-41
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 24, 31));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 32, 35));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 36, 39));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 40, 40));
        //合并42-55
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 41, 50));
        sheet1.addMergedRegion(new CellRangeAddress(1, 1, 51, 52));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 53, 53));
        sheet1.addMergedRegion(new CellRangeAddress(1, 3, 54, 54));


        //第三行
        Row row2 = sheet1.createRow(row1num.getAndIncrement());
        String[] row_second = {"", "", "", "", "", "", "目标达成", "月度实际", "考核事故总数", "工伤事故总数",
                "目标达成", "消防事故总数", "目标达成", "财产存世安全事故总数", "损失额", "目标达成", "损失额",
                "目标达成", "事故起数", "目标达成", "食品安全事故总数", "行政监察次数", "目标达成", "罚款额",
                "目标达成", "安全隐患整改率", "上期总隐患", "上期总整改数", "本期新增隐患", "本期新增整改", "累积总隐患数", "累积总整改数",
                "目标达成", "特种设备定检率", "特种设备数量", "已检合格数量", "目标达成", "特种作业证上岗率", "应持证人数", "实际持证人员",
                "", "目标达成", "新员工", "", "", "在岗员工", "", "", "月度安全活动", "", "", "", "", "", ""};
        for (int i = 0; i < row_second.length; i++) {
            Cell cell = row2.createCell(i);
            cell.setCellValue(row_second[i]);
            cell.setCellStyle(cellStyle);
        }
        //第 7-40行的合并方式相同
        for (int i = 6; i <= 39; i++) {
            sheet1.addMergedRegion(new CellRangeAddress(2, 3, i, i));
        }
        sheet1.addMergedRegion(new CellRangeAddress(2, 3, 41, 41));
        sheet1.addMergedRegion(new CellRangeAddress(2, 2, 42, 44));
        sheet1.addMergedRegion(new CellRangeAddress(2, 2, 45, 47));
        sheet1.addMergedRegion(new CellRangeAddress(2, 2, 48, 50));

        //第四行
        Row row3 = sheet1.createRow(row1num.getAndIncrement());
        String[] row_third = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "培训率", "新员工人", "按期培训人数", "培训率", "在岗人数", "参与培训人数", "集团/事业部",
                "分公司", "仓", "专职安全人员", "兼职安全人员", "", ""};
        for (int i = 0; i < row_third.length; i++) {
            Cell cell = row3.createCell(i);
            cell.setCellValue(row_third[i]);
            cell.setCellStyle(cellStyle);
        }
        //数据行
        //记录一下标题用了多少行
        int title_line = row1num.get();
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.outPutData(title_line, data, sheet1, row1num, cellStyle);

    }

    /**
     * “分公司汇总”的sheet
     * @param workbook
     */
    private void createSheet2(Workbook workbook,List<List> data){
        Sheet sheet2 = workbook.createSheet("分公司汇总");
        AtomicInteger row1num = new AtomicInteger();
        //第一行
        Row row0 = sheet2.createRow(row1num.getAndIncrement());
        Cell title = row0.createCell(0);
        title.setCellValue("百世集团EHS双周报");
        sheet2.addMergedRegion(new CellRangeAddress(0, 0, 0, 54));
        title.setCellStyle(cellStyle);
        //第二行
        Row row1 = sheet2.createRow(row1num.getAndIncrement());
        String[] row_first = {"地区", "分公司", "在职人数", "总部办公人数", "一线人数", "未遂事件数量",
                "考核工伤事故 月度目标0.23‰", "", "", "", "消防事故 月度目标 0", "",
                "财产损失安全事故 月度目标 0", "", "", "环境污染 0", "", "烧车事故", "", "食品安全事故 0（店加）",
                "", "政府检查罚款", "", "", "安全隐患", "", "", "", "", "", "", "",
                "特种设备", "", "", "", "特种作业", "", "", "", "新改扩建场所安全评估",
                "安全培训 新员工安全培训覆盖率（入职7日内） 100%", "", "", "", "", "", "", "", "", "",
                "安全人员数量", "", "工业风扇", "备注"};
        for (int i = 0; i < row_first.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(row_first[i]);
            cell.setCellStyle(cellStyle);
        }
        //合并1-6
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 3, 3));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 4, 4));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 5, 5));
        //合并7-24
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 6, 9));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 10, 11));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 12, 14));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 15, 16));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 17, 18));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 19, 20));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 21, 23));
        //合并25-41
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 24, 31));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 32, 35));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 36, 39));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 40, 40));
        //合并
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 41, 50));
        sheet2.addMergedRegion(new CellRangeAddress(1, 1, 51, 52));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 53, 53));
        sheet2.addMergedRegion(new CellRangeAddress(1, 3, 54, 54));


        //第三行
        Row row2 = sheet2.createRow(row1num.getAndIncrement());
        String[] row_second = {"", "", "", "", "", "", "目标达成", "月度实际", "考核事故总数", "工伤事故总数",
                "目标达成", "消防事故总数", "目标达成", "财产存世安全事故总数", "损失额", "目标达成", "损失额",
                "目标达成", "事故起数", "目标达成", "食品安全事故总数", "行政监察次数", "目标达成", "罚款额",
                "目标达成", "安全隐患整改率", "上期总隐患", "上期总整改数", "本期新增隐患", "本期新增整改", "累积总隐患数", "累积总整改数",
                "目标达成", "特种设备定检率", "特种设备数量", "已检合格数量", "目标达成", "特种作业证上岗率", "应持证人数", "实际持证人员",
                "", "目标达成", "新员工", "", "", "在岗员工", "", "", "月度安全活动", "", "", "", "", "", ""};
        for (int i = 0; i < row_second.length; i++) {
            Cell cell = row2.createCell(i);
            cell.setCellValue(row_second[i]);
            cell.setCellStyle(cellStyle);
        }
        //第 7-40行的合并方式相同
        for (int i = 6; i <= 39; i++) {
            sheet2.addMergedRegion(new CellRangeAddress(2, 3, i, i));
        }
        sheet2.addMergedRegion(new CellRangeAddress(2, 3, 41, 41));
        sheet2.addMergedRegion(new CellRangeAddress(2, 2, 42, 44));
        sheet2.addMergedRegion(new CellRangeAddress(2, 2, 45, 47));
        sheet2.addMergedRegion(new CellRangeAddress(2, 2, 48, 50));

        //第四行
        Row row3 = sheet2.createRow(row1num.getAndIncrement());
        String[] row_third = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "培训率", "新员工人", "按期培训人数", "培训率", "在岗人数", "参与培训人数", "集团/事业部",
                "分公司", "仓", "专职安全人员", "兼职安全人员", "", ""};
        for (int i = 0; i < row_third.length;i++) {
            Cell cell = row3.createCell(i);
            cell.setCellValue(row_third[i]);
            cell.setCellStyle(cellStyle);
        }

        //数据行
        //记录一下标题用了多少行
        int title_line = row1num.get();
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.outPutData(title_line, data, sheet2, row1num, cellStyle);


    }

    /**
     * “快递事业部”的sheet
     * @param workbook
     */
    private void createSheet3(Workbook workbook,List<List> data) {
        Sheet sheet3 = workbook.createSheet("快递事业部");
        AtomicInteger row1num = new AtomicInteger();
        //第一行
        Row row0 = sheet3.createRow(row1num.getAndIncrement());
        Cell title = row0.createCell(0);
        title.setCellValue("快递事业部双周报");
        sheet3.addMergedRegion(new CellRangeAddress(0, 0, 0, 54));
        title.setCellStyle(cellStyle);
        //第二行
        Row row1 = sheet3.createRow(row1num.getAndIncrement());
        String[] row_first = {"分公司", "", "分拨/仓", "在职人数", "职能人数", "一线人数", "未遂事件数量", "考核工伤事故 月度目标0.23‰", "", "", "",
                "消防事故 月度目标 0", "", "财产损失安全事故 月度目标 0", "", "", "环境污染 0", "", "烧车事故 0", "", "食品安全事故 0（店加）", "",
                "政府检查罚款 0", "", "", "安全隐患 整改97%（上期整改截止日在考核周期内的隐患整改数，不是整改数）", "", "", "", "", "", "", "",
                "特种设备 定检率100%", "", "", "", "特种作业 持证上岗率110%", "", "", "", "安全培训 新员工安全培训覆盖率（入职7日内） 100%", "", "", "", "", "",
                "月度安全活动", "安全人员数量", "", "工业风扇", "备注"};
        for (int i = 0; i < row_first.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(row_first[i]);
        }
        //合并1-7
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 0, 1));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 3, 3));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 4, 4));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 5, 5));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 6, 6));
        //合并8-20
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 7, 10));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 11, 12));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 13, 15));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 16, 17));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 18, 19));
        //合并21-41
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 20, 21));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 22, 24));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 25, 32));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 33, 36));
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 37, 40));
        //合并
        sheet3.addMergedRegion(new CellRangeAddress(1, 1, 41, 46));
        sheet3.addMergedRegion(new CellRangeAddress(1, 2, 47, 47));
        sheet3.addMergedRegion(new CellRangeAddress(1, 2, 48, 49));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 50, 50));
        sheet3.addMergedRegion(new CellRangeAddress(1, 3, 51, 51));


        //第三行
        Row row2 = sheet3.createRow(row1num.getAndIncrement());
        String[] row_second = {"", "", "", "", "", "", "", "目标达成", "月度实际", "考核事故总数", "工伤事故总数",
                "目标达成", "消防事故总数（事故类型为火灾）", "目标达成", "财产损失安全事故总数", "损失额（事故报告小计）",
                "目标达成", "起数", "目标达成", "起数", "目标达成", "食品安全事故总数", "行政检察次数",
                "目标达成", "罚款额（来源于子政府触发模块）", "目标达成", "隐患整改率（统计即按时整改率）", "上期总隐患（上期整改预计截止日在上期的考核周期）",
                "上期总数整改完成数（上期整改预计截止日在上期考核周期内完成整改）", "本期新增隐患（本期整改截止日在考核周期内的隐患）", "本期新增整改完成数（本期整改截止日在考核周期内的整改数）",
                "累积总隐患数", "累积总整改数", "目标达成", "特种设备定检率", "特种设备数量（截止到导出时间）", "已检合格数量", "目标达成", "特种作业持证上岗率", "应持证人数", "实际持证人员",
                "目标达成", "新员工", "", "", "在岗员工（含新员工）", "", "", "", "", "", ""};
        for (int i = 0; i < row_second.length; i++) {
            Cell cell = row2.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(row_second[i]);
        }
        //7-41的合并规则相同
        for (int i = 7 ; i <= 41; i++) {
            sheet3.addMergedRegion(new CellRangeAddress(2, 3, i, i));
        }
        sheet3.addMergedRegion(new CellRangeAddress(2, 2, 42, 44));
        sheet3.addMergedRegion(new CellRangeAddress(2, 2, 45, 46));

        //第四行
        Row row3 = sheet3.createRow(3);
        String[] row_third = {"", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "培训率（按期培训人数/一线人数）", "新员工人数（新员工培训这个课程类型下的学员数）",
                "按期培训人数", "培训率", "累积参与培训人数（截止到导出时间）", "活动次数", "专职安全人员", "兼职安全人员", "", ""};
        for (int i = 0; i < row_third.length; i++) {
            Cell cell = row3.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(row_third[i]);
        }
        //数据行
        //记录一下标题用了多少行
        int title_line = row1num.get();
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.outPutData(title_line, data, sheet3, row1num, cellStyle);
    }

    /**
     * “快运事业部”的sheet
     * @param workbook
     */
    private void createSheet4(Workbook workbook,List<List> data) {
        Sheet sheet4 = workbook.createSheet("快运事业部");
        AtomicInteger row1num = new AtomicInteger();
        //第一行
        Row row0 = sheet4.createRow(row1num.getAndIncrement());
        Cell title = row0.createCell(0);
        title.setCellValue("百世集团EHS双周报");
        sheet4.addMergedRegion(new CellRangeAddress(0, 0, 0, 54));
        title.setCellStyle(cellStyle);
        //第二行
        Row row1 = sheet4.createRow(row1num.getAndIncrement());
        String[] row_first = {"地区", "分拨库/分公司", "在职人数", "总部办公人数", "一线人数", "未遂事件数量",
                "考核工伤事故 月度目标0.23‰", "", "", "", "消防事故 月度目标 0", "",
                "财产损失安全事故 月度目标 0", "", "", "环境污染 0", "", "烧车事故", "", "食品安全事故 0（店加）",
                "", "政府检查罚款", "", "", "安全隐患", "", "", "", "", "", "", "",
                "特种设备", "", "", "", "特种作业", "", "", "", "新改扩建场所安全评估",
                "安全培训 新员工安全培训覆盖率（入职7日内） 100%", "", "", "", "", "", "", "", "", "",
                "安全人员数量", "", "工业风扇", "备注"};
        for (int i = 0; i < row_first.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(row_first[i]);
            cell.setCellStyle(cellStyle);
        }
        //合并1-6
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 3, 3));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 4, 4));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 5, 5));
        //合并7-24
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 6, 9));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 10, 11));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 12, 14));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 15, 16));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 17, 18));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 19, 20));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 21, 23));
        //合并25-41
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 24, 31));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 32, 35));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 36, 39));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 40, 40));
        //合并42-55
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 41, 50));
        sheet4.addMergedRegion(new CellRangeAddress(1, 1, 51, 52));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 53, 53));
        sheet4.addMergedRegion(new CellRangeAddress(1, 3, 54, 54));


        //第三行
        Row row2 = sheet4.createRow(row1num.getAndIncrement());
        String[] row_second = {"", "", "", "", "", "", "目标达成", "月度实际", "考核事故总数", "工伤事故总数",
                "目标达成", "消防事故总数", "目标达成", "财产存世安全事故总数", "损失额", "目标达成", "损失额",
                "目标达成", "事故起数", "目标达成", "食品安全事故总数", "行政监察次数", "目标达成", "罚款额",
                "目标达成", "安全隐患整改率", "上期总隐患", "上期总整改数", "本期新增隐患", "本期新增整改", "累积总隐患数", "累积总整改数",
                "目标达成", "特种设备定检率", "特种设备数量", "已检合格数量", "目标达成", "特种作业证上岗率", "应持证人数", "实际持证人员",
                "", "目标达成", "新员工", "", "", "在岗员工", "", "", "月度安全活动", "", "", "", "", "", ""};
        for (int i = 0; i < row_second.length; i++) {
            Cell cell = row2.createCell(i);
            cell.setCellValue(row_second[i]);
            cell.setCellStyle(cellStyle);
        }
        //第 7-40行的合并方式相同
        for (int i = 6; i <= 39; i++) {
            sheet4.addMergedRegion(new CellRangeAddress(2, 3, i, i));
        }
        sheet4.addMergedRegion(new CellRangeAddress(2, 3, 41, 41));
        sheet4.addMergedRegion(new CellRangeAddress(2, 2, 42, 44));
        sheet4.addMergedRegion(new CellRangeAddress(2, 2, 45, 47));
        sheet4.addMergedRegion(new CellRangeAddress(2, 2, 48, 50));

        //第四行
        Row row3 = sheet4.createRow(row1num.get());
        String[] row_third = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "培训率", "新员工人", "按期培训人数", "培训率", "在岗人数", "参与培训人数", "集团/事业部",
                "分公司", "仓", "专职安全人员", "兼职安全人员", "", ""};
        for (int i = 0; i < row_third.length;i++) {
            Cell cell = row3.createCell(i);
            cell.setCellValue(row_third[i]);
            cell.setCellStyle(cellStyle);
        }
        //数据行
        //记录一下标题用了多少行
        int title_line = row1num.get();
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.outPutData(title_line, data, sheet4, row1num, cellStyle);
    }

    /**
     * “供应链事业部”的sheet
     * @param workbook
     */
    private void createSheet5(Workbook workbook,List<List> data) {
        Sheet sheet5 = workbook.createSheet("供应链事业部");
        AtomicInteger row1num = new AtomicInteger();
        //第一行
        Row row0 = sheet5.createRow(row1num.getAndIncrement());
        Cell title = row0.createCell(0);
        title.setCellValue("百世集团EHS双周报");
        sheet5.addMergedRegion(new CellRangeAddress(0, 0, 0, 54));
        title.setCellStyle(cellStyle);
        //第二行
        Row row1 = sheet5.createRow(row1num.getAndIncrement());
        String[] row_first = {"地区", "仓库/分公司", "在职人数", "总部办公人数", "一线人数", "未遂事件数量",
                "考核工伤事故 月度目标0.23‰", "", "", "", "消防事故 月度目标 0", "",
                "财产损失安全事故 月度目标 0", "", "", "环境污染 0", "", "烧车事故", "", "食品安全事故 0（店加）",
                "", "政府检查罚款", "", "", "安全隐患", "", "", "", "", "", "", "",
                "特种设备", "", "", "", "特种作业", "", "", "", "新改扩建场所安全评估",
                "安全培训 新员工安全培训覆盖率（入职7日内） 100%", "", "", "", "", "", "", "", "", "",
                "安全人员数量", "", "工业风扇", "备注"};
        for (int i = 0; i < row_first.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(row_first[i]);
            cell.setCellStyle(cellStyle);
        }
        //合并1-6
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 3, 3));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 4, 4));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 5, 5));
        //合并7-24
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 6, 9));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 10, 11));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 12, 14));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 15, 16));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 17, 18));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 19, 20));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 21, 23));
        //合并25-41
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 24, 31));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 32, 35));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 36, 39));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 40, 40));
        //合并42-55
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 41, 50));
        sheet5.addMergedRegion(new CellRangeAddress(1, 1, 51, 52));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 53, 53));
        sheet5.addMergedRegion(new CellRangeAddress(1, 3, 54, 54));


        //第三行
        Row row2 = sheet5.createRow(row1num.getAndIncrement());
        String[] row_second = {"", "", "", "", "", "", "目标达成", "月度实际", "考核事故总数", "工伤事故总数",
                "目标达成", "消防事故总数", "目标达成", "财产存世安全事故总数", "损失额", "目标达成", "损失额",
                "目标达成", "事故起数", "目标达成", "食品安全事故总数", "行政监察次数", "目标达成", "罚款额",
                "目标达成", "安全隐患整改率", "上期总隐患", "上期总整改数", "本期新增隐患", "本期新增整改", "累积总隐患数", "累积总整改数",
                "目标达成", "特种设备定检率", "特种设备数量", "已检合格数量", "目标达成", "特种作业证上岗率", "应持证人数", "实际持证人员",
                "", "目标达成", "新员工", "", "", "在岗员工", "", "", "月度安全活动", "", "", "", "", "", ""};
        for (int i = 0; i < row_second.length; i++) {
            Cell cell = row2.createCell(i);
            cell.setCellValue(row_second[i]);
            cell.setCellStyle(cellStyle);
        }
        //第 7-40行的合并方式相同
        for (int i = 6; i <= 39; i++) {
            sheet5.addMergedRegion(new CellRangeAddress(2, 3, i, i));
        }
        sheet5.addMergedRegion(new CellRangeAddress(2, 3, 41, 41));
        sheet5.addMergedRegion(new CellRangeAddress(2, 2, 42, 44));
        sheet5.addMergedRegion(new CellRangeAddress(2, 2, 45, 47));
        sheet5.addMergedRegion(new CellRangeAddress(2, 2, 48, 50));

        //第四行
        Row row3 = sheet5.createRow(row1num.get());
        String[] row_third = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "培训率", "新员工人", "按期培训人数", "培训率", "在岗人数", "参与培训人数", "集团/事业部",
                "分公司", "仓", "专职安全人员", "兼职安全人员", "", ""};
        for (int i = 0; i < row_third.length;i++) {
            Cell cell = row3.createCell(i);
            cell.setCellValue(row_third[i]);
            cell.setCellStyle(cellStyle);
        }
        //数据行
        //记录一下标题用了多少行
        int title_line = row1num.get();
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.outPutData(title_line, data, sheet5, row1num, cellStyle);
    }

    /**
     * “店加事业部”的sheet
     * @param workbook
     */
    private void createSheet6(Workbook workbook,List<List> data) {
        Sheet sheet6 = workbook.createSheet("店加事业部");
        AtomicInteger row1num = new AtomicInteger();
        //第一行
        Row row0 = sheet6.createRow(row1num.getAndIncrement());
        Cell title = row0.createCell(0);
        title.setCellValue("百世集团EHS双周报");
        sheet6.addMergedRegion(new CellRangeAddress(0, 0, 0, 54));
        title.setCellStyle(cellStyle);
        //第二行
        Row row1 = sheet6.createRow(row1num.getAndIncrement());
        String[] row_first = {"地区", "仓库/分公司", "在职人数", "总部办公人数", "一线人数", "未遂事件数量",
                "考核工伤事故 月度目标0.23‰", "", "", "", "消防事故 月度目标 0", "",
                "财产损失安全事故 月度目标 0", "", "", "环境污染 0", "", "烧车事故", "", "食品安全事故 0（店加）",
                "", "政府检查罚款", "", "", "安全隐患", "", "", "", "", "", "", "",
                "特种设备", "", "", "", "特种作业", "", "", "", "新改扩建场所安全评估",
                "安全培训 新员工安全培训覆盖率（入职7日内） 100%", "", "", "", "", "", "", "", "", "",
                "安全人员数量", "", "工业风扇", "备注"};
        for (int i = 0; i < row_first.length; i++) {
            Cell cell = row1.createCell(i);
            cell.setCellValue(row_first[i]);
            cell.setCellStyle(cellStyle);
        }
        //合并1-6
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 3, 3));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 4, 4));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 5, 5));
        //合并7-24
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 6, 9));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 10, 11));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 12, 14));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 15, 16));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 17, 18));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 19, 20));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 21, 23));
        //合并25-41
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 24, 31));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 32, 35));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 36, 39));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 40, 40));
        //合并42-55
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 41, 50));
        sheet6.addMergedRegion(new CellRangeAddress(1, 1, 51, 52));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 53, 53));
        sheet6.addMergedRegion(new CellRangeAddress(1, 3, 54, 54));


        //第三行
        Row row2 = sheet6.createRow(row1num.getAndIncrement());
        String[] row_second = {"", "", "", "", "", "", "目标达成", "月度实际", "考核事故总数", "工伤事故总数",
                "目标达成", "消防事故总数", "目标达成", "财产存世安全事故总数", "损失额", "目标达成", "损失额",
                "目标达成", "事故起数", "目标达成", "食品安全事故总数", "行政监察次数", "目标达成", "罚款额",
                "目标达成", "安全隐患整改率", "上期总隐患", "上期总整改数", "本期新增隐患", "本期新增整改", "累积总隐患数", "累积总整改数",
                "目标达成", "特种设备定检率", "特种设备数量", "已检合格数量", "目标达成", "特种作业证上岗率", "应持证人数", "实际持证人员",
                "", "目标达成", "新员工", "", "", "在岗员工", "", "", "月度安全活动", "", "", "", "", "", ""};
        for (int i = 0; i < row_second.length; i++) {
            Cell cell = row2.createCell(i);
            cell.setCellValue(row_second[i]);
            cell.setCellStyle(cellStyle);
        }
        //第 7-40行的合并方式相同
        for (int i = 6; i <= 39; i++) {
            sheet6.addMergedRegion(new CellRangeAddress(2, 3, i, i));
        }
        sheet6.addMergedRegion(new CellRangeAddress(2, 3, 41, 41));
        sheet6.addMergedRegion(new CellRangeAddress(2, 2, 42, 44));
        sheet6.addMergedRegion(new CellRangeAddress(2, 2, 45, 47));
        sheet6.addMergedRegion(new CellRangeAddress(2, 2, 48, 50));

        //第四行
        Row row3 = sheet6.createRow(row1num.get());
        String[] row_third = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                "培训率", "新员工人", "按期培训人数", "培训率", "在岗人数", "参与培训人数", "集团/事业部",
                "分公司", "仓", "专职安全人员", "兼职安全人员", "", ""};
        for (int i = 0; i < row_third.length;i++) {
            Cell cell = row3.createCell(i);
            cell.setCellValue(row_third[i]);
            cell.setCellStyle(cellStyle);
        }
        //数据行
        //记录一下标题用了多少行
        int title_line = row1num.get();
        ExcelUtils excelUtils = new ExcelUtils();
        excelUtils.outPutData(title_line, data, sheet6, row1num, cellStyle);
    }
}
