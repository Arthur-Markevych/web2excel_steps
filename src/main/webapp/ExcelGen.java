package main.webapp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.util.IOUtils;


import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelGen {


    private static String path = "C:\\Users\\Diveloper\\Documents\\excelPOIwriteTest\\";

    private static String sheetName = "test sheet";

    public static void write() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        //Set columns sizes
        int colNum = 0;
        sheet.setColumnWidth(colNum++, 40 * 256); // Photo column
        sheet.setColumnWidth(colNum++, 20 * 256); // Name column
        sheet.setColumnWidth(colNum++, 20 * 256); // Options column
        sheet.setColumnWidth(colNum++, 15 * 256); // Amount column
        sheet.setColumnWidth(colNum++, 17 * 256); // Total Price column
        sheet.setColumnWidth(colNum++, 14 * 256); // Delivery column


        // Name of set Header
        int rowCount = 4;
        XSSFRow rowHead = sheet.createRow(rowCount++);
        rowHead.setHeightInPoints(21f);
        XSSFCell headCell = rowHead.createCell(0);

        setMerge(sheet, rowHead.getRowNum(), rowHead.getRowNum(), 0, 5, true);
        headCell.setCellValue("test header".toUpperCase());

        XSSFCellStyle style1 = wb.createCellStyle();
        style1.setAlignment(HorizontalAlignment.CENTER);
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        style1.setBorderLeft(BorderStyle.MEDIUM);
        style1.setBorderRight(BorderStyle.MEDIUM);
        style1.setBorderTop(BorderStyle.MEDIUM);
        style1.setBorderBottom(BorderStyle.MEDIUM);

        style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(204, 204, 204)));
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        //Font style
        Font font = wb.createFont();
//        font.setColor(IndexedColors.BLACK.getIndex());
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style1.setFont(font);
        headCell.setCellStyle(style1);

        //Main info Header Cell Style
        XSSFCellStyle summaryCellStyle = wb.createCellStyle();
        summaryCellStyle.cloneStyleFrom(style1); // clone style properties
        summaryCellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(238, 238, 238)));

        // Header Product Items
        XSSFRow itemsHeader = sheet.createRow(rowCount++); // items header row
        XSSFCellStyle itHeaderCellStyle = wb.createCellStyle();
        itHeaderCellStyle.cloneStyleFrom(summaryCellStyle);

        int itCellNum = 0;

        XSSFCell imgHeaderCell = itemsHeader.createCell(itCellNum++); // image cell header
        imgHeaderCell.setCellStyle(itHeaderCellStyle);
        imgHeaderCell.setCellValue("photo".toUpperCase());
        // Fill the rest of order items header
        setItemsHeaderCell("name".toUpperCase(), itemsHeader, itCellNum++, itHeaderCellStyle);
        setItemsHeaderCell("options".toUpperCase(),itemsHeader, itCellNum++, itHeaderCellStyle);
        setItemsHeaderCell("amount".toUpperCase(),itemsHeader, itCellNum++, itHeaderCellStyle);
        setItemsHeaderCell("total price".toUpperCase(),itemsHeader, itCellNum++, itHeaderCellStyle);
        setItemsHeaderCell("delivery".toUpperCase(),itemsHeader, itCellNum++, itHeaderCellStyle);

        // Product cell style
        XSSFCellStyle productCellStyle = wb.createCellStyle();
        productCellStyle.setAlignment(HorizontalAlignment.CENTER);
        productCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // Product description row --- --- --- -- forEach() -- --- --- --- ---
        int productRowNum = rowCount++;
        int productRowNumEnd = productRowNum + 4;
        XSSFRow productRow = sheet.createRow(productRowNum);
        XSSFRow pOneRow = sheet.createRow(rowCount++);
        XSSFRow pTwoRow = sheet.createRow(rowCount++);
        XSSFRow pThreeRow = sheet.createRow(rowCount++);
        XSSFRow pFourRow = sheet.createRow(rowCount++);
        setProductRowsStyle(productCellStyle, 21f, productRow, pOneRow, pTwoRow, pThreeRow, pFourRow);

        // --add image
        addImage(wb, sheet, "C:\\Users\\Diveloper\\IdeaProjects\\web2excel\\test_files\\1.png");
        // --/add image

        CellRangeAddress imgMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 0,0);
//        sheet.addMergedRegion(imgMerge);
        int productCell = 0;
        XSSFCell imgCell = productRow.createCell(productCell++);


        // Cells merges
        CellRangeAddress optionsMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 2,2);
        CellRangeAddress amountMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 3,3);
        CellRangeAddress totalPriceMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 4,4);
        CellRangeAddress deliveryPriceMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 5,5);
        addMergeRegions(sheet, imgMerge, optionsMerge, amountMerge, totalPriceMerge, deliveryPriceMerge);
//        int pColnum = 2;
//        setMerge(sheet, productRowNum, productRowNumEnd, pColnum, pColnum++, true);
//        setMerge(sheet, productRowNum, productRowNumEnd, pColnum, pColnum++, true);
//        setMerge(sheet, productRowNum, productRowNumEnd, pColnum, pColnum++, true);
//        setMerge(sheet, productRowNum, productRowNumEnd, pColnum, pColnum++, true);


//        rowCount += 4;


        // Summary rows ---------------------------------------------

        // Item's price row
        XSSFRow itemPriceRow = sheet.createRow(rowCount++);
        itemPriceRow.setHeightInPoints(20f);
        XSSFCell itLeftCell = itemPriceRow.createCell(0);
        XSSFCell itRightCell = itemPriceRow.createCell(4);
        int prRowNum = itemPriceRow.getRowNum(); // get current row num
        setMerge(sheet, prRowNum, prRowNum, 0, 3, true);
        setMerge(sheet, prRowNum, prRowNum, 4, 5, true);
        itLeftCell.setCellStyle(summaryCellStyle);
        itRightCell.setCellStyle(summaryCellStyle);
        itLeftCell.setCellValue("Item's price: ");
        itRightCell.setCellValue(0);

        // Delivery row
        XSSFRow deliveryRow = sheet.createRow(rowCount++);
        XSSFCell delLeftCell = deliveryRow.createCell(0);
        XSSFCell delRightCell = deliveryRow.createCell(4);
        int delRowNum = deliveryRow.getRowNum();

        setMerge(sheet, delRowNum, delRowNum, 0, 3, true);
        setMerge(sheet, delRowNum, delRowNum, 4, 5, true);
        delLeftCell.setCellStyle(summaryCellStyle);
        delRightCell.setCellStyle(summaryCellStyle);
        delLeftCell.setCellValue("Delivery price:");
        delRightCell.setCellValue(0);

        // Total price row
        XSSFRow totalPriceRow = sheet.createRow(rowCount++);
        XSSFCell tplLeftCell = totalPriceRow.createCell(0);
        XSSFCell tplRightCell = totalPriceRow.createCell(4);
        int tpRowNum = totalPriceRow.getRowNum();

        setMerge(sheet, tpRowNum, tpRowNum, 0, 3, true);
        setMerge(sheet, tpRowNum, tpRowNum, 4, 5, true);
        tplLeftCell.setCellStyle(summaryCellStyle);
        tplRightCell.setCellStyle(summaryCellStyle);
        tplLeftCell.setCellValue("Total price:");
        tplRightCell.setCellValue(0);

        // Write to file
        try (FileOutputStream out = new FileOutputStream(path + "t1.xlsx")) {
            wb.write(out);
            out.flush();
            wb.close();
        }

    }

    protected static void addImage(Workbook wb, Sheet sheet, String imgPath) throws IOException {
        try (InputStream inputStream = new FileInputStream(imgPath)) {
            byte[] bytes = IOUtils.toByteArray(inputStream);
            int picIndex = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

            CreationHelper helper = wb.getCreationHelper();
            //Creates the top-level drawing patriarch.
            Drawing drawing = sheet.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();

            //create an anchor with upper left cell _and_ bottom right cell
            anchor.setCol1(0); //Column B
            anchor.setRow1(7); //Row 3
            anchor.setCol2(1); //Column C
            anchor.setRow2(11); //Row 4

            //Creates a picture
            Picture pict = drawing.createPicture(anchor, picIndex);

            //Reset the image to the original size
            //pict.resize(); //don't do that. Let the anchor resize the image!

            //Create the Cell B3
            Cell cell = sheet.createRow(2).createCell(1);

            //set width to n character widths = count characters * 256
            //int widthUnits = 20*256;
            //sheet.setColumnWidth(1, widthUnits);

            //set height to n points in twips = n * 20
            //short heightUnits = 60*20;
            //cell.getRow().setHeight(heightUnits);

        }
    }

    protected static void setMerge(Sheet sheet, int numRow, int untilRow, int numCol, int untilCol, boolean border) {
        CellRangeAddress cellMerge = new CellRangeAddress(numRow, untilRow, numCol, untilCol);
        sheet.addMergedRegion(cellMerge);
        if (border) {
            setBordersToMergedCells(sheet, cellMerge);
        }
    }

    protected static void setProductRowsStyle(XSSFCellStyle style, float height, XSSFRow... rows) {
        for (XSSFRow row : rows) {
            row.setHeightInPoints(height);
            row.setRowStyle(style);
        }
    }

    protected static void setBordersToMergedCells(Sheet sheet, CellRangeAddress rangeAddress) { // Borders
        RegionUtil.setBorderTop(BorderStyle.MEDIUM, rangeAddress, sheet);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM, rangeAddress, sheet);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM, rangeAddress, sheet);
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM, rangeAddress, sheet);
    }

    protected static void setItemsHeaderCell(String value, XSSFRow row, int cellNum, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(cellNum);
        cell.setCellStyle(style);
        cell.setCellValue(value);
    }

    protected static void addMergeRegions(Sheet sheet, CellRangeAddress... margeAddress) {
        for (CellRangeAddress rangeAddress : margeAddress) {
            sheet.addMergedRegion(rangeAddress);
//            setBordersToMergedCells(sheet, rangeAddress); // todo fix problem with adding borders
        }
    }

    protected static void setSummuryCellsStyle(XSSFRow row, int leftCol, int rightCol, XSSFCellStyle style) {
        int rowNum = row.getRowNum();
        //todo think how to reuse code of setting Cell Styles
    }

    public static void main(String[] args) throws IOException, ParseException {
        write();
        System.out.println("Done! " + new SimpleDateFormat("HH:mm:ss").format(new Date()));
    }
}
