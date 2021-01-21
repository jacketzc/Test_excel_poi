package com.jac.model;

/**
 * @author ：jacketzc
 * Created in 2020/11/11 14:29
 * 单元格的合并信息
 */
public class DataCellMergeInfo {
    private int firstRow;
    private int lastRow;
    private int firstCol;
    private int lastCol;

    public DataCellMergeInfo() {
    }

    public DataCellMergeInfo(int firstRow, int lastRow, int firstCol, int lastCol) {
        this.firstRow = firstRow;
        this.lastRow = lastRow;
        this.firstCol = firstCol;
        this.lastCol = lastCol;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public int getLastCol() {
        return lastCol;
    }

    @Override
    public String toString() {
        return "DataCellMergeInfo{" +
                "firstRow=" + firstRow +
                ", lastRow=" + lastRow +
                ", firstCol=" + firstCol +
                ", lastCol=" + lastCol +
                '}';
    }
}
