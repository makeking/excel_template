package com.wangy.myapplication.poi;

import android.util.Log;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class PoiCopyUtils {


    /**
     * 复制一个单元格样式到目的单元格样式
     *
     * @param fromStyle
     * @param toStyle
     */
    public static void copyCellStyle(XSSFCellStyle fromStyle,
                                     XSSFCellStyle toStyle) {
        toStyle.setAlignment(fromStyle.getAlignment());
        //边框和边框颜色
        toStyle.setBorderBottom(fromStyle.getBorderBottom());
        toStyle.setBorderLeft(fromStyle.getBorderLeft());
        toStyle.setBorderRight(fromStyle.getBorderRight());
        toStyle.setBorderTop(fromStyle.getBorderTop());
        toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
        toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
        toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
        toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());

        //背景和前景
        toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
        toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());

        toStyle.setDataFormat(fromStyle.getDataFormat());
        toStyle.setFillPattern(fromStyle.getFillPattern());
//		toStyle.setFont(fromStyle.getFont(null));
        toStyle.setHidden(fromStyle.getHidden());
        toStyle.setIndention(fromStyle.getIndention());//首行缩进
        toStyle.setLocked(fromStyle.getLocked());
        toStyle.setRotation(fromStyle.getRotation());//旋转
        toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());
        toStyle.setWrapText(fromStyle.getWrapText());

    }

    /**
     * Sheet复制
     *
     * @param fromSheet
     * @param toSheet
     * @param copyValueFlag
     */
    public static void copySheet(XSSFWorkbook wb, XSSFSheet fromSheet, XSSFSheet toSheet,
                                 boolean copyValueFlag, List<CellRangeAddress> rangeList, XSSFSheet sheet) {
        //合并区域处理
        mergerRegion(fromSheet, toSheet);
        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext(); ) {
            XSSFRow tmpRow = (XSSFRow) rowIt.next();
            XSSFRow newRow = toSheet.createRow(tmpRow.getRowNum());
            //行复制
            copyRow(wb, tmpRow, newRow, copyValueFlag, rangeList, sheet);
        }
    }

    /**
     * 行复制功能
     *
     * @param fromRow
     * @param toRow
     * @param rangeList
     * @param sheet
     */
    public static void copyRow(XSSFWorkbook wb, XSSFRow fromRow, XSSFRow toRow, boolean copyValueFlag, List<CellRangeAddress> rangeList, XSSFSheet sheet) {
        Iterator<Cell> it = fromRow.cellIterator();
        while ( it.hasNext() ) {
            XSSFCell tmpCell = (XSSFCell) it.next();
            XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
//            Log.e("tag"," 当前的数字是 " + fromRow.getRowNum() );
            copyCell(wb, tmpCell, newCell, copyValueFlag, rangeList, sheet, tmpCell.getColumnIndex(), fromRow.getRowNum());
        }

    }


    /**
     * 复制原有sheet的合并单元格到新创建的sheet
     *
     * @param fromSheet 新创建sheet
     * @param toSheet   原有的sheet
     */
    public static void mergerRegion(XSSFSheet fromSheet, XSSFSheet toSheet) {
        int sheetMergerCount = fromSheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress address = fromSheet.getMergedRegion(i);
            toSheet.addMergedRegion(address);
        }
    }

    /**
     * 复制单元格
     *
     * @param srcCell
     * @param distCell
     * @param copyValueFlag true则连同cell的内容一起复制
     * @param rangeList
     * @param sheet
     * @param cellNum
     */
    public static void copyCell(XSSFWorkbook wb, XSSFCell srcCell, XSSFCell distCell,
                                boolean copyValueFlag, List<CellRangeAddress> rangeList, XSSFSheet sheet, int cellNum, int rowNum) {
        try {
            Map<String, Object> combineCell = isCombineCell(rangeList, srcCell);
            if ((Boolean) combineCell.get("flag")) {
                // 判断是否是合并单元格
                Log.e("tag", "当前选中的cell 是合并单元格 ");
                // 获取合并行和合并列
                int mergedRow = (int) combineCell.get("mergedRow");
                int mergedCol = (int) combineCell.get("mergedCol");
                // todo 合并单元格

                // 行号 2- 3
                // 列号 0-0
                // 4 5  4
                // 6 7  4
                // 8 9
                // 10 11
//                CellRangeAddress region = new CellRangeAddress(rowNum * mergedRow, rowNum * mergedRow +1 , 0, 0 );
//                CellRangeAddress region = new CellRangeAddress(rowNum, rowNum + mergedRow - 1, cellNum, cellNum + mergedCol - 1);
//                CellRangeAddress region = sheet.getMergedRegion(0);
//                sheet.addMergedRegion(region);
            }
            XSSFCellStyle newstyle = wb.createCellStyle();
            copyCellStyle(srcCell.getCellStyle(), newstyle);
//        distCell.setEncoding(srcCell.getEncoding());
            //样式
            distCell.setCellStyle(newstyle);
            //评论
            if (srcCell.getCellComment() != null) {
                distCell.setCellComment(srcCell.getCellComment());
            }
            // 不同数据类型处理
            int srcCellType = srcCell.getCellType();
            distCell.setCellType(srcCellType);
            if (copyValueFlag) {
                if (srcCellType == XSSFCell.CELL_TYPE_NUMERIC) {
                    if (HSSFDateUtil.isCellDateFormatted(srcCell)) {
                        distCell.setCellValue(srcCell.getDateCellValue());
                    } else {
                        distCell.setCellValue(srcCell.getNumericCellValue());
                    }
                } else if (srcCellType == XSSFCell.CELL_TYPE_STRING) {
                    distCell.setCellValue(srcCell.getRichStringCellValue());
                } else if (srcCellType == XSSFCell.CELL_TYPE_BLANK) {
                    // nothing21
                } else if (srcCellType == XSSFCell.CELL_TYPE_BOOLEAN) {
                    distCell.setCellValue(srcCell.getBooleanCellValue());
                } else if (srcCellType == XSSFCell.CELL_TYPE_ERROR) {
                    distCell.setCellErrorValue(srcCell.getErrorCellValue());
                } else if (srcCellType == XSSFCell.CELL_TYPE_FORMULA) {
                    distCell.setCellFormula(srcCell.getCellFormula());
                } else { // nothing29
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static Map<String, Object> isCombineCell(List<CellRangeAddress> listCombineCell, Cell cell)
            throws Exception {
        int firstC = 0;
        int lastC = 0;
        int firstR = 0;
        int lastR = 0;
        String cellValue = null;
        Boolean flag = false;
        int mergedRow = 0;
        int mergedCol = 0;
        Map<String, Object> result = new HashMap<>();
        result.put("flag", flag);
        for (CellRangeAddress ca : listCombineCell) {
            //获得合并单元格的起始行, 结束行, 起始列, 结束列
            firstC = ca.getFirstColumn();
            lastC = ca.getLastColumn();
            firstR = ca.getFirstRow();
            lastR = ca.getLastRow();
            //判断cell是否在合并区域之内，在的话返回true和合并行列数
            if (cell!= null){
                if (cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR) {
                    if (cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC) {
                        flag = true;
                        mergedRow = lastR - firstR + 1;
                        mergedCol = lastC - firstC + 1;
                        result.put("flag", true);
                        result.put("mergedRow", mergedRow);
                        result.put("mergedCol", mergedCol);
                        break;
                    }
                }
            }
        }
        return result;
    }

}
