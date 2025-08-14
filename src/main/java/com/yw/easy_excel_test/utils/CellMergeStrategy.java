package com.yw.easy_excel_test.utils;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.Set;

/**
 * 自定义单元格合并策略
 */
public class CellMergeStrategy extends AbstractMergeStrategy {

    /**
     * 唯一标识列(该列cell内容一定是全局唯一的,当cell的值相等时才能进行下一列合并)
     */
    private Integer uniqueColumnIndex;
    /**
     * 从哪一行开始合并
     */
    private Integer mergeRowIndex = 0;
    /**
     * 合并列编号,从0开始
     */
    private List<Integer> mergeColumnIndex = Lists.newArrayList();

    private CellMergeStrategy() {
    }

    public CellMergeStrategy(Integer uniqueColumnIndex, Integer mergeRowIndex, Set<Integer> mergeColumnIndex) {
        mergeColumnIndex.stream().forEach(item -> {
            this.mergeColumnIndex.add(item);
        });
        this.mergeColumnIndex.stream().sorted();

        if (null == uniqueColumnIndex) {
            this.uniqueColumnIndex = this.mergeColumnIndex.get(0);
        } else {
            this.uniqueColumnIndex = uniqueColumnIndex;
        }
        this.mergeRowIndex = mergeRowIndex;
    }

    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer relativeRowIndex) {
        int curColIndex = cell.getColumnIndex();
        int curRowIndex = cell.getRowIndex();

        // 判断该列是否需要合并
        if (!mergeColumnIndex.contains(curColIndex)) {
            return;
        }

        if (curRowIndex > mergeRowIndex) {
            for (int i = 0; i < mergeColumnIndex.size(); i++) {
                if (curColIndex == mergeColumnIndex.get(i)) {
                    this.mergeRow(sheet, cell, curRowIndex, curColIndex, uniqueColumnIndex);
                    break;
                }
            }
        }
    }


    /**
     * 向上合并单元格
     *
     * @param cell           当前单元格
     * @param rowIndex       当前行
     * @param colIndex       当前列
     * @param uniqueColIndex 唯一标识列(该列cell内容一定是全局唯一的,当cell的值相等时才能进行下一列合并)
     */
    private void mergeRow(Sheet sheet, Cell cell, int rowIndex, int colIndex, int uniqueColIndex) {
        Object curCellValue = getCellValue(cell);
        Object preCellValue = getCellValue(cell.getSheet().getRow(rowIndex - 1).getCell(colIndex));
        boolean cellEqual = preCellValue.equals(curCellValue);

        boolean baseCellEqual = true;
        if (colIndex >= uniqueColIndex) {
            Object baseCellValue = getCellValue(cell.getRow().getCell(uniqueColIndex));
            Object preBaseCellValue = getCellValue(cell.getSheet().getRow(rowIndex - 1).getCell(uniqueColIndex));
            baseCellEqual = baseCellValue.equals(preBaseCellValue);
        }

        /**
         * 合并条件
         * 1. 将当前单元格数据与上一个单元格数据比较,相同则执行合并逻辑
         * 2. 唯一标识列内容相同，才能进行下一列合并
         */
        if (!(cellEqual && baseCellEqual)) {
            return;
        }

        List<CellRangeAddress> mergeRegionList = sheet.getMergedRegions();
        boolean isMerged = false;
        for (int i = 0; i < mergeRegionList.size() && !isMerged; i++) {
            CellRangeAddress cellRange = mergeRegionList.get(i);
            // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
            if (cellRange.isInRange(rowIndex - 1, colIndex)) {
                sheet.removeMergedRegion(i);
                cellRange.setLastRow(rowIndex);
                sheet.addMergedRegion(cellRange);
                isMerged = true;
            }
        }
        // 若上一个单元格未被合并，则新增合并单元
        if (!isMerged) {
            CellRangeAddress cellRange = new CellRangeAddress(rowIndex - 1, rowIndex, colIndex, colIndex);
            sheet.addMergedRegion(cellRange);
        }
    }

    private Object getCellValue(Cell baseCell) {
        return CellType.STRING.equals(baseCell.getCellType()) ? baseCell.getStringCellValue() : baseCell.getNumericCellValue();
    }

}
