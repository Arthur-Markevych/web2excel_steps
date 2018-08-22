package main.webapp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.util.IOUtils;


import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelGen {

    private final static int NAME_COL_ROW_NUM = 5;

//    private static String path = "C:\\Users\\Diveloper\\Documents\\excelPOIwriteTest\\";
    private static String path = "C:\\Users\\Artur_Markevych\\Documents\\excel_test\\";

    private static String sheetName = "test sheet";

    private static String imgPath = "C:\\Users\\Artur_Markevych\\Documents\\Pre_Prod_java_q3q4_2018\\task1 – git pracrice I\\images2\\006.jpg";

    public static void write() throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);

        //Set columns sizes
        int colNum = 0;
        sheet.setColumnWidth(colNum++, 40 * 256); // Photo column
        sheet.setColumnWidth(colNum++, 25 * 256); // Name column
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

        XSSFCellStyle leftRightBsAlCenter = wb.createCellStyle(); // Name cell style
        leftRightBsAlCenter.cloneStyleFrom(productCellStyle);
        leftRightBsAlCenter.setBorderLeft(BorderStyle.MEDIUM);
        leftRightBsAlCenter.setBorderRight(BorderStyle.MEDIUM);

        // merged cell style
        XSSFCellStyle bigCellStyle = wb.createCellStyle();
        bigCellStyle.cloneStyleFrom(leftRightBsAlCenter);
        bigCellStyle.setBorderTop(BorderStyle.MEDIUM);
        bigCellStyle.setBorderBottom(BorderStyle.MEDIUM);

        // first name cell style
        XSSFCellStyle topBorderStyle = wb.createCellStyle();
        topBorderStyle.cloneStyleFrom(leftRightBsAlCenter);
        topBorderStyle.setBorderTop(BorderStyle.MEDIUM);

        // Product description row --- --- --- -- forEach() -- --- --- --- ---
        // ----- ----- ----- ----- ----- ----- ----- ----- ----- ----- ---- ----
        for (TestModel m : TestModel.getAll()) {
            int productRowNum = rowCount++;
            int productRowNumEnd = productRowNum + 4;
            XSSFRow productRow = sheet.createRow(productRowNum);
            XSSFRow pTwoRow = sheet.createRow(rowCount++);
            XSSFRow pThreeRow = sheet.createRow(rowCount++);
            XSSFRow pFourRow = sheet.createRow(rowCount++);
            XSSFRow pFiveRow = sheet.createRow(rowCount++);
            setProductRowsStyle(productCellStyle, 24f, productRow, pTwoRow, pThreeRow, pFourRow, pFiveRow);

            // --add image
//            addImage(wb, sheet, productRow, m.getImgPath());
            // --/add image

            addImageInCell(sheet, m.getImgPath(), sheet.createDrawingPatriarch(), 0, productRowNum);

            CellRangeAddress imgMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 0, 0);
//        sheet.addMergedRegion(imgMerge);
            // Name cells
            int productCell = 0;
            XSSFCell imgCell = productRow.createCell(productCell++);
            XSSFCell nameCell = productRow.createCell(productCell);
            XSSFCell heightCell = pTwoRow.createCell(productCell);
            XSSFCell widthCell = pThreeRow.createCell(productCell);
            XSSFCell depthCell = pFourRow.createCell(productCell);
            XSSFCell volumeCell = pFiveRow.createCell(productCell++);

            nameCell.setCellStyle(topBorderStyle);
            heightCell.setCellStyle(leftRightBsAlCenter);
            widthCell.setCellStyle(leftRightBsAlCenter);
            depthCell.setCellStyle(leftRightBsAlCenter);
            volumeCell.setCellStyle(leftRightBsAlCenter);

            nameCell.setCellValue(m.getName().toUpperCase());
            heightCell.setCellValue("height: " + m.getHeight());
            widthCell.setCellValue("width: " + m.getWidth());
            depthCell.setCellValue("depth: " + m.getDepth());
            volumeCell.setCellValue("volume: " + m.getVolume());


            XSSFCell optionsCell = productRow.createCell(productCell++);
            optionsCell.setCellStyle(bigCellStyle);
            optionsCell.setCellValue(m.getOptions());

            XSSFCell amountCell = productRow.createCell(productCell++);
            amountCell.setCellStyle(bigCellStyle);
            amountCell.setCellValue(m.getAmount() + " psc.");

            XSSFCell priceCell = productRow.createCell(productCell++);
            priceCell.setCellStyle(bigCellStyle);
            priceCell.setCellValue(m.getPrice() * m.getAmount());

            XSSFCell deliveryCell = productRow.createCell(productCell++);
            deliveryCell.setCellStyle(bigCellStyle);
            deliveryCell.setCellValue("no defined");



            // Cells merges
            CellRangeAddress optionsMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 2, 2);
            CellRangeAddress amountMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 3, 3);
            CellRangeAddress totalPriceMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 4, 4);
            CellRangeAddress deliveryPriceMerge = new CellRangeAddress(productRowNum, productRowNumEnd, 5, 5);
            addMergeRegions(sheet, imgMerge, optionsMerge, amountMerge, totalPriceMerge, deliveryPriceMerge);
        }


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
        try (FileOutputStream out = new FileOutputStream(path + "test.xlsx")) {
            wb.write(out);
            out.flush();
            wb.close();
        }

    }

    protected static void addImage(Workbook wb, Sheet sheet, XSSFRow row, String imgPath) throws IOException {
        try (InputStream inputStream = new FileInputStream(imgPath)) {

        // ------------
//            Image img = ImageIO.read(inputStream);
//            BufferedImage tempJPG = resizeImage(img, 100, 100);
        // ------------


            byte[] bytes = IOUtils.toByteArray(inputStream);
            int picIndex = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);

            CreationHelper helper = wb.getCreationHelper();
            //Creates the top-level drawing patriarch.
            Drawing drawing = sheet.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();

            //create an anchor with upper left cell _and_ bottom right cell
            anchor.setCol1(0);
            anchor.setRow1(row.getRowNum());
            anchor.setCol2(1);
            anchor.setRow2(row.getRowNum() + NAME_COL_ROW_NUM);

//            anchor.setDx1(0);
//            anchor.setDy1(0);
//            anchor.setDx1((int) (0.5 * sheet.getColumnWidthInPixels(0) * Units.EMU_PER_PIXEL));
//            anchor.setDx2((int) (0.4 * ImageUtils.getRowHeightInPixels(sheet, row.getRowNum()) * Units.EMU_PER_PIXEL));

            int colWidth = sheet.getColumnWidth(0);
            int rowHeigth = row.getHeight();


            //Creates a picture
            Picture pict = drawing.createPicture(anchor, picIndex);
//            pict.resize();
//            pict.resize( 0.2);
//            int picWidth = pict.getImageDimension().width;
//            double resize = (double) colWidth / 10_000;
//            pict.resize(resize, resize);
            //set height to n points in twips = n * 20
//            short heightUnits = (short) (picHeight * 20);
//            row.setHeight(heightUnits);

            //Reset the image to the original size
//            pict.resize(); //don't do that. Let the anchor resize the image!
        }
    }
    // temporary, to delete. Just try.
    protected static void addImageInCell(Sheet sheet, String url, Drawing<?> drawing, int colNumber, int rowNumber) throws IOException {
        InputStream inputStream = new FileInputStream(url);
        BufferedImage imageIO = ImageIO.read(inputStream);
        int height = imageIO.getHeight();
        int width = imageIO.getWidth();
        int relativeHeight = (int) (((double) height / width) * 28.5);
        new AddDimensionedImage().addImageToSheet(colNumber, rowNumber, sheet, drawing, new URL(url), 30, relativeHeight,
                AddDimensionedImage.EXPAND_ROW_AND_COLUMN);

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

        for (int i = 0; i < margeAddress.length; i++) {
            sheet.addMergedRegion(margeAddress[i]);
            setBordersToMergedCells(sheet, margeAddress[i]);
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

    /**
     * This function resize the image file and returns the BufferedImage object that can be saved to file system.
     */
    public static BufferedImage resizeImage(final Image image, int width, int height) {
        final BufferedImage bufferedImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
        final Graphics2D graphics2D = bufferedImage.createGraphics();
        graphics2D.setComposite(AlphaComposite.Src);
        //below three lines are for RenderingHints for better image quality at cost of higher processing time
        graphics2D.setRenderingHint(RenderingHints.KEY_INTERPOLATION,RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        graphics2D.setRenderingHint(RenderingHints.KEY_RENDERING,RenderingHints.VALUE_RENDER_QUALITY);
        graphics2D.setRenderingHint(RenderingHints.KEY_ANTIALIASING,RenderingHints.VALUE_ANTIALIAS_ON);
        graphics2D.drawImage(image, 0, 0, width, height, null);
        graphics2D.dispose();
        return bufferedImage;
    }

}
