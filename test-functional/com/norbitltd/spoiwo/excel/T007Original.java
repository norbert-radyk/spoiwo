package com.norbitltd.spoiwo.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class T007Original {

    public void fillsAndColors() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow((short) 1);

        // Aqua background
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(CellStyle.BIG_SPOTS);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Orange "foreground", foreground being the fill foreground not the font color.
        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell = row.createCell((short) 2);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

}
