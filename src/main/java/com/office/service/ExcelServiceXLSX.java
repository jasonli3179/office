package com.office.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;

/**
 * xlsx新格式Excel操作
 * XLSX格式使用不带前缀的工具类
 */
@Service
public class ExcelServiceXLSX {

    /**
     * 在指定单元格处向下插入一列
     *
     * @param sheet
     * @param startRow  开始插入的行数
     * @param endColumn 有效信息最大列数
     */
    public void move(Sheet sheet, int startRow, int endColumn) {
        //遍历所有行,从第3行开始
        for (int x = startRow; x < sheet.getLastRowNum() + 1; x++) {
            //当前行
            Row row = sheet.getRow(x);

            //遍历行中每个单元格
            for (int y = endColumn; y > 0; y--) {
                Cell newCell = row.getCell(y) != null ? row.getCell(y) : row.createCell(y);
                Cell oldCell = row.getCell(y - 1) != null ? row.getCell(y - 1) : row.createCell(y - 1);

                //复制单元格样式
//                newCell.getCellStyle().cloneStyleFrom(oldCell.getCellStyle());
                newCell.setCellStyle(oldCell.getCellStyle());
                //设置单元格宽
                sheet.setColumnWidth(newCell.getColumnIndex(), sheet.getColumnWidth(endColumn - 1));
                //单元格赋值（字符串格式）
                oldCell.setCellType(CellType.STRING);
                newCell.setCellType(CellType.STRING);
                newCell.setCellValue(oldCell.getStringCellValue());

            }
        }

    }

    /**
     * 从指定单元格开始向下插入一列数据
     *
     * @param sheet
     * @param column 列号
     * @param value
     */
    public void insertColumn(Sheet sheet, int column, String value) {
        //循环插入列column，值为value
        for (int x = 3; x < sheet.getLastRowNum() + 1; x++) {
            Row row = sheet.getRow(x);
            Cell cell = row.getCell(column);
            cell.setCellValue(value);
        }
        //新插入列的列头信息
    }

    /**
     * 指定单元格赋值
     * @param sheet
     * @param row 行数
     * @param column 列数
     * @param value 值
     */
    public void insert(Sheet sheet, int row, int column,String value) {
        Cell cell = sheet.getRow(row).getCell(column);
        cell.setCellValue(value);
    }
}
