import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * http://poi.apache.org/spreadsheet/quick-guide.html#NewWorkbook
 *
 */
public class TryDifferent {


    public static void main(String[] args) throws Exception {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Java to Excel");
        sheet.setColumnWidth(0, 10000);


        //add picture data to this workbook.
        InputStream is = new FileInputStream("Image1.jpeg");
        byte[] bytes = IOUtils.toByteArray(is);
        int pictureIndex = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        is.close();

        CreationHelper helper = workbook.getCreationHelper();

        // Create the drawing patriarch.  This is the top level container for all shapes.
        Drawing drawing = sheet.createDrawingPatriarch();

        //add a picture shape
        ClientAnchor anchor = helper.createClientAnchor();
        //set top-left corner of the picture,
        //subsequent call of Picture#resize() will operate relative to it
        anchor.setCol1(0);
        anchor.setRow1(0);
        Picture picture = drawing.createPicture(anchor, pictureIndex);

        //auto-size picture relative to its top-left corner
        picture.resize(1, 1);



//        sheet.addMergedRegion(new CellRangeAddress(0, 4, 0, 3));

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(1);
        cell.getRow().setHeightInPoints(150);

        cell.setCellValue("Some text");
        Font font = getFont(workbook);
        CellStyle style = getCellStyle(workbook);
        style.setFont(font);

        cell.setCellStyle(style);


        writeExcel(workbook, "five");
    }


    public static CellStyle getCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return  style;
    }

    public static Font getFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setColor(IndexedColors.YELLOW.getIndex());
        font.setBold(true);
        font.setItalic(true);
        font.setFontHeightInPoints((short)16);
        font.setFontName("Helvetia");
        return font;
    }

    public static void  addCell(Row row, String value, int n) {
        Cell cell = row.createCell(n);
        cell.setCellValue(value);
    }

    public static void writeExcel(Workbook workbook, String fname) {

        fname += ".xls";
        if(workbook instanceof XSSFWorkbook) fname += "x";
        try (OutputStream out = new FileOutputStream(fname)) {
            workbook.write(out);
        } catch (FileNotFoundException e) {
            System.out.println(e.getMessage());
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }

    }



}
