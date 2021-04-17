package com.example.word.utils;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

public class ExcelUtil {
    private static XSSFWorkbook newExcelCreat = new XSSFWorkbook();
    private static XSSFCellStyle newstyle = newExcelCreat.createCellStyle();
    public static void main(String[] args) {
        long time1 = System.currentTimeMillis();
        String path = "D:\\ceshi\\excel\\新建文件夹";
        List<String> list = Arrays.asList(new File(path).listFiles()).stream().map(x -> x.getPath()).collect(Collectors.toList());
        mergexcel(list,"merge.xlsx",path);
        long time2 = System.currentTimeMillis();
        System.out.println(time2 - time1);
    }

    /**
     * * 合并多个ExcelSheet
     *
     * @param files 文件字符串(file.toString)集合,按顺序进行合并，合并的Excel中Sheet名称不可重复
     * @param excelName 合并后Excel名称(包含后缀.xslx)
     * @param dirPath 存储目录
     * @return
     * @Date: 2020/9/18 15:31
     */
    public static void mergexcel(List<String> files, String excelName, String dirPath) {

        // 遍历每个源excel文件，TmpList为源文件的名称集合
        for (String fromExcelName : files) {
            try (InputStream in = new FileInputStream(fromExcelName)) {
                XSSFWorkbook fromExcel = new XSSFWorkbook(in);
                int length = fromExcel.getNumberOfSheets();
                // 长度为1时
                if (length <= 1) {
                    XSSFSheet oldSheet = fromExcel.getSheetAt(0);
                    XSSFSheet newSheet = newExcelCreat.createSheet(oldSheet.getSheetName());
                    copySheet(newExcelCreat, oldSheet, newSheet);
                } else {
                    // 遍历每个sheet
                    for (int i = 0; i < length; i++) {
                        XSSFSheet oldSheet = fromExcel.getSheetAt(i);
                        XSSFSheet newSheet = newExcelCreat.createSheet(oldSheet.getSheetName());
                        copySheet(newExcelCreat, oldSheet, newSheet);
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        // 定义新生成的xlxs表格文件
        String allFileName = dirPath + File.separator + excelName;
        try (FileOutputStream fileOut = new FileOutputStream(allFileName)) {
            newExcelCreat.write(fileOut);
            fileOut.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                newExcelCreat.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 合并单元格
     *
     * @param fromSheet
     * @param toSheet
     */
    private static void mergeSheetAllRegion(XSSFSheet fromSheet, XSSFSheet toSheet) {
        int num = fromSheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress cellR = fromSheet.getMergedRegion(i);
            try {
                toSheet.addMergedRegion(cellR);
            } catch (IllegalArgumentException e) {
                continue;
            }

        }
    }

    /**
     * 复制单元格
     *
     * @param wb
     * @param fromCell
     * @param toCell
     */
    private static void copyCell(XSSFWorkbook wb, XSSFCell fromCell, XSSFCell toCell) {

        // 复制单元格样式
        newstyle.cloneStyleFrom(fromCell.getCellStyle());
        // 样式
        toCell.setCellStyle(newstyle);
        if (fromCell.getCellComment() != null) {
            toCell.setCellComment(fromCell.getCellComment());
        }
        // 不同数据类型处理
        CellType fromCellType = CellType.forInt(fromCell.getCellType());
        toCell.setCellType(fromCellType);
        if (fromCellType == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(fromCell)) {
                toCell.setCellValue(fromCell.getDateCellValue());
            } else {
                toCell.setCellValue(fromCell.getNumericCellValue());
            }
        } else if (fromCellType == CellType.STRING) {
            toCell.setCellValue(fromCell.getRichStringCellValue());
        } else if (fromCellType == CellType.BLANK) {
            // nothing21
        } else if (fromCellType == CellType.BOOLEAN) {
            toCell.setCellValue(fromCell.getBooleanCellValue());
        } else if (fromCellType == CellType.ERROR) {
            toCell.setCellErrorValue(fromCell.getErrorCellValue());
        } else if (fromCellType == CellType.FORMULA) {
            toCell.setCellFormula(fromCell.getCellFormula());
        } else {
            // nothing29
        }
    }

    /**
     * 行复制功能
     *
     * @param wb
     * @param oldRow
     * @param toRow
     */
    private static void copyRow(XSSFWorkbook wb, XSSFRow oldRow, XSSFRow toRow) {
        toRow.setHeight(oldRow.getHeight());
        for (Iterator cellIt = oldRow.cellIterator(); cellIt.hasNext(); ) {
            XSSFCell tmpCell = (XSSFCell) cellIt.next();
            XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
            copyCell(wb, tmpCell, newCell);
        }
    }

    /**
     * Sheet复制
     *
     * @param wb
     * @param fromSheet
     * @param toSheet
     */
    private static void copySheet(XSSFWorkbook wb, XSSFSheet fromSheet, XSSFSheet toSheet) {
        mergeSheetAllRegion(fromSheet, toSheet);
        // 设置列宽
        int length = fromSheet.getRow(fromSheet.getFirstRowNum()).getLastCellNum();
        for (int i = 0; i <= length; i++) {
            toSheet.setColumnWidth(i, fromSheet.getColumnWidth(i));
        }
        for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext(); ) {
            XSSFRow oldRow = (XSSFRow) rowIt.next();
            XSSFRow newRow = toSheet.createRow(oldRow.getRowNum());
            copyRow(wb, oldRow, newRow);
        }
    }
}
