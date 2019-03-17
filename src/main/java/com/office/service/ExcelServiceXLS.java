package com.office.service;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

/**
 * xls旧格式Excel操作
 * xls格式使用HSSF开头的工具类
 */
@Service
public class ExcelServiceXLS {


    /**
     * 在指定单元格处向下插入一列
     *
     * @param sheet
     * @param startRow 开始插入的行数
     * @param endColumn 有效信息最大列数
     */
    public void move(HSSFSheet sheet, int startRow, int endColumn) {
        //遍历所有行,从第3行开始
        for (int x = startRow; x < sheet.getLastRowNum() + 1; x++) {
            //当前行
            HSSFRow row = sheet.getRow(x);
            if (row == null) {
                continue;
            }

            //遍历行中每个单元格
            for (int y = endColumn; y > 0; y--) {
                HSSFCell newCell = row.getCell(y) != null ? row.getCell(y) : row.createCell(y);
                HSSFCell oldCell = row.getCell(y - 1) != null ? row.getCell(y - 1) : row.createCell(y - 1);

                //复制单元格样式，(单元格的宽高不包含在样式中)
//                newCell.getCellStyle().cloneStyleFrom(oldCell.getCellStyle());
                newCell.setCellStyle(oldCell.getCellStyle());
                //设置单元格宽
                sheet.setColumnWidth(newCell.getColumnIndex(), sheet.getColumnWidth(endColumn - 1));
                //单元格赋值
                oldCell.setCellType(CellType.STRING);
                newCell.setCellType(CellType.STRING);
                newCell.setCellValue(oldCell.getStringCellValue());

            }
        }

    }


    /**
     * 指定列插入数据
     *
     * @param sheet
     * @param column
     * @param value
     */
    public void insertColumn(HSSFSheet sheet, int column, String value) {
        for (int x = 3; x < sheet.getLastRowNum() + 1; x++) {
            HSSFRow row = sheet.getRow(x);
            if (row == null) {
                continue;
            }
            HSSFCell cell = row.getCell(column);
            cell.setCellValue(value);
        }
    }

    /**
     * 指定单元格赋值
     * @param sheet
     * @param row 行数
     * @param column 列数
     * @param value 值
     */
    public void insert(Sheet sheet, int row, int column, String value) {
        Cell cell = sheet.getRow(row).getCell(column);
        cell.setCellValue(value);
    }
}
