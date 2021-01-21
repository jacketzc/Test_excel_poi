package com.jac.service;

import com.jac.exception.CreateExcelException;
import com.jac.model.DataCell;
import com.jac.model.DataCellMergeInfo;
import com.jac.model.ExcelType;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.CollectionUtils;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;


/**
 * @author ：jacketzc
 * @date ：Created in 2020/11/10 9:16
 */
@Slf4j
public class ExcelUtils {
    //合并操作默认探索的深度
    public final int DEFAULT_STEP = 2;


    /**
     * 将普通的单元格元素转换为 DataCell元素
     * 所有初始的 list必须经过该方法处理，否则抛出异常
     * @param list
     * @return
     */
    public List<List> convertToDataCellList(List<List> list) {
        List<List> res = new ArrayList<>();
        list.forEach(l->{res.add((List) l.stream().map(DataCell::convertToDataCell).collect(Collectors.toList()));});
        return res;
    }

    /**
     * 为某个sheet中单元格添加合并信息
     * 合并信息只包含上下的合并信息！（在通常的业务中，数据单元格并没有出现过左右合并的现象）
     * 为了节约资源，合并操作的默认深度为 2 ，即list中index为2的元素不会再和下一个进行比较
     *
     * @param list
     */
    public List<List> addMergeInfo(String sheetName, List<List> list) {
        return addMergeInfo(sheetName, list, DEFAULT_STEP);
    }

    /**
     * 为单元格添加合并信息
     *
     * @param list
     * @param step 探索的深度
     */
    public List<List> addMergeInfo(String sheetName, List<List> list, int step) {
        //如果list中没有数据，那就直接返回
        if (CollectionUtils.isEmpty(list)) {
            log.info("sheet" + sheetName + "中没有数据");
            return null;
        }


        //如果不是处理过的 list，将其处理
        if (!isDataCellList(list)) list = convertToDataCellList(list);

        //找到需要添加合并信息的cell，并且添加合并信息
        //先将list转换为数组，便于操作
        int size = list.get(0).size();
        Object[][] arrayData = new Object[list.size()][size];
        for (int i = 0; i < list.size(); i++) {
            arrayData[i] = list.get(i).toArray();
        }
        // 对每个深度的单元格信息进行合并
        // 这里采用了lazy模式，不论上级是否在同一个已经合并的单元格中，在当前深度都会对所有的单元格进行尝试合并
        for (int col = 0; col < step; col++) {
            for (int row = 0; row < list.size(); row++) {
                int last = checkNext(arrayData, (DataCell) arrayData[row][col], row, col);
                //如果当前行在下方有相同的元素，则为当前单元格元素添加合并信息
                if (last != row) {
                    DataCell dataCell = (DataCell) list.get(row).get(col);
                    dataCell.setMergeInfo(new DataCellMergeInfo(row, last, col, col));
                    //不应该重复合并
                    row = last - 1;
                }

            }
        }
        return list;
    }

    /**
     * 判断list中的元素是否已经被转换为 DataCell
     * @param list
     * @return
     */
    private boolean isDataCellList(List<List> list) {
        boolean flag = list.get(0).get(0) instanceof DataCell;
        return flag;
    }

    /**
     * 递归函数：
     * 查找与对比元素相同的最后一个元素
     * @param strs
     * @param now
     * @return 最后一个元素的数组下标
     */
    private int checkNext(Object[][] strs, DataCell now ,int row ,int col) {
        //如果到最后一个元素，直接返回
        if (row==strs.length-1) return row;

        if (now.equals(strs[row + 1][col])) {
            return checkNext(strs, (DataCell) strs[row + 1][col], row + 1, col);
        }
        return row;
    }


    public void outPutData(int title_line, List<List> data, Sheet sheet, AtomicInteger row1num, CellStyle cellStyle) {
        // 空集合就没必要操作了
        if (CollectionUtils.isEmpty(data))  return;
        data.forEach(d->{
            Row row = sheet.createRow(row1num.getAndIncrement());
            for (int i = 0; i < d.size(); i++) {
                DataCell dataCell = (DataCell) d.get(i);
                Cell cell = row.createCell(i);
                cell.setCellValue(dataCell.getCellInfo());
                cell.setCellStyle(cellStyle);

                //如果该元素有合并信息，则添加合并信息
                DataCellMergeInfo mergeInfo = dataCell.getMergeInfo();
                if (mergeInfo != null) {
                    // mergeInfo 中记录的 row信息时不包括头部信息的，所以要添上去
                    sheet.addMergedRegion(new CellRangeAddress(mergeInfo.getFirstRow()+title_line , mergeInfo.getLastRow()+title_line , mergeInfo.getFirstCol(), mergeInfo.getLastCol()));
                }
            }
        });

    }
    /**
     * 创建一个空的、xlsx格式的 Workbook
     * @return
     * @throws Exception
     */
    public  Workbook createEmptyXlsxWorkbook() throws Exception{
        return createEmptyWorkbook(ExcelType.XLSX);
    }
    /**
     * 创建一个空的、xlsx格式的 Workbook
     * @return
     * @throws Exception
     */
    public  Workbook createEmptyXlsWorkbook() throws Exception{
        return createEmptyWorkbook(ExcelType.XLS);
    }
    /**
     * 创建一个空的Workbook
     * @return
     * @throws Exception
     */
    public  Workbook createEmptyWorkbook(ExcelType type) throws Exception{
        // 根据type类型生成不同的Workbook
        Workbook workbook = null;
        switch (type) {
            case XLS:
                workbook = (Workbook) Class.forName("org.apache.poi.hssf.usermodel.HSSFWorkbook").newInstance();
                break;
            case XLSX: workbook = (Workbook) Class.forName("org.apache.poi.xssf.usermodel.XSSFWorkbook").newInstance();
                break;
            default: throw new CreateExcelException("指定的待生成文件类型错误！可选的类型有：ExcelType.XLS、ExcelType.XLSX");
        }
        return workbook;
    }

}
