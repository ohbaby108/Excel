package com.xinrui;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.OutputStream;

public class StudentInfo {
    public static void main(String[] args)throws Exception {
        // 创建一个Excel文件
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建一个工作表
        HSSFSheet sheet = workbook.createSheet("学生表一");
        // 标题行合并单元格
        CellRangeAddress region = new CellRangeAddress(
                0, // first row
                0, // last row
                0, // first column
                4 // last column
        );
        sheet.addMergedRegion(region);
        HSSFRow hssfRow = sheet.createRow(0);

        HSSFCell headCell = hssfRow.createCell(0);
        // 确定标题的
        headCell.setCellValue("学生信息表");

        // 设置单元格格式居中
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        headCell.setCellStyle(cellStyle);

        // 添加表头行
        hssfRow = sheet.createRow(1);// 第几行 行的位置编号
        // 添加表头内容
        headCell = hssfRow.createCell(0);
        headCell.setCellValue("姓名");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(1);
        headCell.setCellValue("年龄");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(2);
        headCell.setCellValue("年级");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(3);
        headCell.setCellValue("分数");
        headCell.setCellStyle(cellStyle);

        headCell = hssfRow.createCell(4);
        headCell.setCellValue("地址");
        headCell.setCellStyle(cellStyle);
        for (int i = 0; i < 100; i++) {//添加100条记录
            //创建表体行
            hssfRow = sheet.createRow(2+i);
            // 创建单元格，并设置值
            HSSFCell cell = hssfRow.createCell(0);
            cell.setCellValue("name");
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(1);
            cell.setCellValue(10);
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(2);
            cell.setCellValue("一年级");
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(3);
            cell.setCellValue("100");
            cell.setCellStyle(cellStyle);

            cell = hssfRow.createCell(4);
            cell.setCellValue("北京大兴");
            cell.setCellStyle(cellStyle);
        }






        // 保存Excel文件
        OutputStream outputStream = new FileOutputStream("C:\\Users\\l\\Desktop\\student.xls");
        workbook.write(outputStream);
        outputStream.close();
    }
}
