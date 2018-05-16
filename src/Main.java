/*import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;*/

public class Main /*extends HttpServlet*/{
/*
        private static final String htmlData = "<html><head><title>Jsoup html parse table</title></head><body><table class=\"tableData\" border=\"0\"><tr>"
                + "<th>Sr.No.</th><th>Studenth>City</th><th>Phone No</th></tr><tr><td>1</td><td>Dixit</td>"
                + "<td>Ahmedabad</td><td>9825098025</td></tr><tr><td>1</td><td>Saharsh</td><td>Ahmedabad</td><td>9825098015</td></tr></table>"
                + "</body></html>";

    @Override
    protected void doGet(HttpServletRequest req, HttpServletResponse response) throws ServletException, IOException {
        Date now = new Date();
        String currDate = new SimpleDateFormat("dd_MM_yyyy_HH_mm").format(now);
        String fileName = "sampleExcel" + currDate;

        // Create book
        HSSFWorkbook wb = new HSSFWorkbook();

        // create excel sheet for page 1
        HSSFSheet sheet = wb.createSheet();

        //Set Header Font
        HSSFFont headerFont = wb.createFont();
        headerFont.setBoldweight(headerFont.BOLDWEIGHT_BOLD);
        headerFont.setFontHeightInPoints((short) 12);

        //Set Header Style
        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        headerStyle.setAlignment(headerStyle.ALIGN_CENTER);
        headerStyle.setFont(headerFont);
        headerStyle.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
        headerStyle.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
        int rowCount = 0;
        Row header;

        */
/*
            Following code parse html table
        *//*

        Document doc = Jsoup.parse(htmlData);

        */
/* Display list of headers for
        tag here i tried to fetch data with class = tableData in table tag
        you can fetch using id or other attribute
        rowCount variable to create row for excel sheet
        *//*

        Cell cell;
        for (Element table : doc.select("table[class=tableData]")) {
            rowCount++;
            // loop through all tr of table
            for (Element row : table.select("tr")) {
                // create row for each tag
                header = sheet.createRow(rowCount);
                // loop through all tag of tag
                Elements ths = row.select("th");
                int count = 0;
                for (Element element : ths) {
                    // set header style
                    cell = header.createCell(count);
                    cell.setCellValue(element.text());
                    cell.setCellStyle(headerStyle);
                    count++;
                }
                // now loop through all td tag
                Elements tds = row.select("td:not([rowspan])");
                count = 0;
                for (Element element : tds) {
                    // create cell for each tag
                    cell = header.createCell(count);
                    cell.setCellValue(element.text());
                    count++;
                }
                rowCount++;
                // set auto size column for excel sheet
                sheet = wb.getSheetAt(0);
                for (int j = 0; j < row.select("th").size(); j++) {
                    sheet.autoSizeColumn(j);
                }
            }
        }
        ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
        wb.write(outByteStream);
        byte[] outArray = outByteStream.toByteArray();

        response.setContentType("application/ms-excel");
        response.setContentLength(outArray.length);
        response.setHeader("Expires:", "0"); // eliminates browser caching
        response.setHeader("Content-Disposition", "attachment; filename=" + fileName + ".xls");
        OutputStream outStream = response.getOutputStream();
        outStream.write(outArray);
        outStream.flush();
        outStream.close();
//        fos.flush();
//        fos.close();
//        outputStream.flush();
    }

    public HttpServletResponse createExcel(HttpServletResponse response) throws IOException {
        return null;
    }
*/

}
