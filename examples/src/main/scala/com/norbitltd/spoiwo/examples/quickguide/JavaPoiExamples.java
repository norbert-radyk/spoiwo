package com.norbitltd.spoiwo.examples.quickguide;

import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

@SuppressWarnings("unused")
class JavaPoiExamples {

    public static void newWorkbook() throws IOException {
        Workbook wb = new XSSFWorkbook();

        wb.getCreationHelper();

        FileOutputStream fileOut = new FileOutputStream("workbook.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }
    public static void newSheet() throws IOException {
        Workbook wb = new XSSFWorkbook();  // or new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); // returns " O'Brien's sales   "
        Sheet sheet3 = wb.createSheet(safeName);
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static void creatingCells() throws IOException {
        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow((short) 0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static void creatingDateCells() throws IOException {
        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow(0);

        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());

        // we style the second cell as a date (and time).  It is important to
        // create a new cell style from the workbook otherwise you can end up
        // modifying the built in style and effecting not only this cell but other cells.
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);

        //you can also set date as java.util.Calendar
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static void workingWithDifferentTypesOfCells() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow((short) 2);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        row.createCell(5).setCellType(CellType.ERROR);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    public static void variousAlignmentOptions() throws Exception {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow((short) 2);
        row.setHeightInPoints(30);

        createCell(wb, row, (short) 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
        createCell(wb, row, (short) 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
        createCell(wb, row, (short) 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        createCell(wb, row, (short) 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
        createCell(wb, row, (short) 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
        createCell(wb, row, (short) 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
        createCell(wb, row, (short) 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("xssf-align.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void createCell(Workbook wb, Row row, short column, HorizontalAlignment halign, VerticalAlignment valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }

    private static void workingWithBorders() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);

        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue(4);

        // Style the cell with borders all around.
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style.setBorderTop(BorderStyle.MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void fillsAndColors() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow((short) 1);

        // Aqua background
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Orange "foreground", foreground being the fill foreground not the font color.
        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell((short) 2);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void mergingCells() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Row row = sheet.createRow((short) 1);
        Cell cell = row.createCell((short) 1);
        cell.setCellValue("This is a test of merging");

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)
                1, //last row  (0-based)
                1, //first column (0-based)
                2  //last column  (0-based)
        ));

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void workingWithFonts() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);

        // Create a new font and alter it.
        Font font = wb.createFont();
        font.setFontHeightInPoints((short) 24);
        font.setFontName("Courier New");
        font.setItalic(true);
        font.setStrikeout(true);

        // Fonts are set into a style so create a new one to use.
        CellStyle style = wb.createCellStyle();
        style.setFont(font);

        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of fonts");
        cell.setCellStyle(style);

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void workingWithMultipleFonts() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("new sheet");

        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        style.setFont(font);
        for (int i = 0; i < 10000; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell((short) 0);
            cell.setCellStyle(style);
        }
    }

    private static void customColors() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell( 0);
        cell.setCellValue("custom XSSF colors");

        XSSFCellStyle style1 = wb.createCellStyle();
        style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 0, 128), new DefaultIndexedColorMap()));
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style1);
    }

    private static void usingNewlinesInCells() throws IOException {
        Workbook wb = new XSSFWorkbook();   //or new HSSFWorkbook();
        Sheet sheet = wb.createSheet();

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("Use \n with word wrap on to create a new line");

        //to enable newlines you need set a cell styles with wrap=true
        CellStyle cs = wb.createCellStyle();
        cs.setWrapText(true);
        cell.setCellStyle(cs);

        //increase row height to accommodate two lines of text
        row.setHeightInPoints((2*sheet.getDefaultRowHeightInPoints()));

        //adjust column width to fit the content
        sheet.autoSizeColumn((short)2);

        FileOutputStream fileOut = new FileOutputStream("ooxml-newlines.xlsx");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void dataFormats() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        CellStyle style;
        DataFormat format = wb.createDataFormat();
        Row row;
        Cell cell;
        short rowNum = 0;
        short colNum = 0;

        row = sheet.createRow(rowNum++);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("0.0"));
        cell.setCellStyle(style);

        row = sheet.createRow(rowNum);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("#,##0.0000"));
        cell.setCellStyle(style);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void fitSheetToOnePage() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        PrintSetup ps = sheet.getPrintSetup();
        sheet.setAutobreaks(true);
        ps.setFitHeight((short)1);
        ps.setFitWidth((short)1);

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void setPrintArea() throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        //sets the print area for the first sheet
        wb.setPrintArea(0, "$A$1:$C$2");

        //Alternatively:
        wb.setPrintArea(
                0, //sheet index
                0, //start column
                1, //end column
                0, //start row
                0  //end row
        );

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void setPageNumbersOnFooter() throws IOException {
        Workbook wb = new XSSFWorkbook(); // or new XSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        Footer footer = sheet.getFooter();

        footer.setRight( "Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages() );

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void setSheetAsSelected() {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("row sheet");
        sheet.setSelected(true);
    }

    private static void setZoomMagnification() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        sheet1.setZoom(75);
    }

    private static void splitAndFreezePanes() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        Sheet sheet3 = wb.createSheet("third sheet");
        Sheet sheet4 = wb.createSheet("fourth sheet");

        // Freeze just one row
        sheet1.createFreezePane( 0, 1, 0, 1 );
        // Freeze just one column
        sheet2.createFreezePane( 1, 0, 1, 0 );
        // Freeze the columns and rows (forget about scrolling position of the lower right quadrant).
        sheet3.createFreezePane( 2, 2 );
        // Create a split with the lower left side being the active quadrant
        sheet4.createSplitPane( 2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT );

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void repeatingRowsAndColumns() throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("Sheet1");
        Sheet sheet2 = wb.createSheet("Sheet2");

        // Set the rows to repeat from row 4 to 5 on the first sheet.
        sheet1.setRepeatingRows(CellRangeAddress.valueOf("4:5"));
        // Set the columns to repeat from column A to C on the second sheet
        sheet2.setRepeatingColumns(CellRangeAddress.valueOf("A:C"));

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }

    private static void headersAndFooters() throws IOException {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        Header header = sheet.getHeader();
        header.setCenter("Center Header");
        header.setLeft("Left Header");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");

        FileOutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        fileOut.close();

    }

}
