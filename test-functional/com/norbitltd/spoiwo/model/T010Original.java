package com.norbitltd.spoiwo.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.*;

public class T010Original {

    public void customColors() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell( 0);
        cell.setCellValue("custom XSSF colors");

        XSSFCellStyle style1 = wb.createCellStyle();
        style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 0, 128)));
        style1.setFillPattern(CellStyle.SOLID_FOREGROUND);
    }
}
