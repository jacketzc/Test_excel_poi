package com.jac.model;

import java.util.Objects;

/**
 * @author ：jacketzc
 * Created in 2020/11/11 14:27
 * 一个单元格中包含的值，以及位置合并信息
 */
public class DataCell {
    //单元格的值，所有类型都将转换为 String类型
    private String cellInfo;
    //单元格的合并信息
    private DataCellMergeInfo mergeInfo;

    public DataCell() {
    }
    public DataCell(String cellInfo) {
        this.cellInfo = cellInfo;
    }

    @Override
    public String toString() {
        return "DataCell{" +
                "cellInfo='" + cellInfo + '\'' +
                ", mergeInfo=" + mergeInfo +
                '}';
    }

    /**
     * 重写的equals方法认为两个 DataCell的 CellInfo相同即为相等
     * @param o
     * @return
     */
    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        DataCell dataCell = (DataCell) o;
        return Objects.equals(cellInfo, dataCell.cellInfo);
    }

    @Override
    public int hashCode() {
        return Objects.hash(cellInfo);
    }

    /**
     * 将普通的单元格元素转换为 DataCell
     * @param cellInfo 单元格元素只能为
     * @return
     */
    public static DataCell convertToDataCell(Object cellInfo) {

        if (cellInfo instanceof String) return new DataCell((String) cellInfo);
//        <? extends String>

        else return new DataCell(String.valueOf(cellInfo));
    }

    public String getCellInfo() {
        return cellInfo;
    }


    public DataCellMergeInfo getMergeInfo() {
        return mergeInfo;
    }

    public void setMergeInfo(DataCellMergeInfo mergeInfo) {
        this.mergeInfo = mergeInfo;
    }
}
