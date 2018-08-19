import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;

public class Main /*extends HttpServlet*/{

    public static final String ENCODING = "UTF-8";
    public static final String OUT_PATH = "imgs\\";


    public static String getHTMLData(String url) throws IOException, FileNotFoundException {
        try (BufferedReader in = new BufferedReader(new InputStreamReader(new FileInputStream(url), ENCODING))) {
            StringBuilder sb = new StringBuilder();
            String str = "";
            while ((str = in.readLine()) != null) {
                sb.append(str + "\n");
            }
        return sb.toString();
        }
    }

    public static void main(String[] args) throws IOException {
        parseHTML(getHTMLData("test.html"));
    }

    public static void  parseHTML(String html) throws IOException {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet = wb.createSheet();

        //Set Header Font
        Font headerFont = wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Calibri");
        headerFont.setFontHeightInPoints((short) 12);

        //Common Cell Font
        Font commonFont = wb.createFont();
        commonFont.setFontName("Calibri");
        commonFont.setFontHeightInPoints((short) 12);

        //Set Header Style
        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFillBackgroundColor(IndexedColors.BLACK.getIndex());
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFont(headerFont);
        headerStyle.setBorderTop(BorderStyle.MEDIUM);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);

        //Set Common SellStyle
        CellStyle commonCellStyle = wb.createCellStyle();
        commonCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        commonCellStyle.setAlignment(HorizontalAlignment.CENTER);
        commonCellStyle.setFont(commonFont);

        int rowCount = 0;
        Row header;
        Cell cell;

        Document doc = Jsoup.parse(html);
        for (Element table : doc.select("table[id=testTable]")) {
            rowCount++;
            for (Element row : table.select("tr")) {
                header = sheet.createRow(rowCount);
                Elements ths = row.select("th");
                int count = 0;
                //Loop through Headers
                for (Element element : ths) {
                    // set header style
                    cell = header.createCell(count);
                    cell.setCellValue(element.text().toUpperCase());
                    cell.setCellStyle(headerStyle);
                    cell.getRow().setHeightInPoints((char) 24);
                    count++;
                }
                // now loop through all td tag
                Elements tds = row.select("td:not([rowspan])");
                count = 0;
                for (Element elementRow : tds) {
                    if (elementRow.childNodeSize() > 0) {
                    Node img = elementRow.selectFirst("img");
                        if (img != null) {
                            String imgPath = img.attr("src");
                            if (imgPath.endsWith(".jpg") || imgPath.endsWith(".png"))
                                saveToFile(imgPath, OUT_PATH + getImgName(imgPath) );
//                        System.out.println("img > " + img.attr("src"));
                        }
                    }

                // create cell for each tag
                    cell = header.createCell(count);
                    cell.getRow().setHeightInPoints(100);
                    cell.setCellStyle(commonCellStyle);
                    cell.setCellValue(elementRow.text());
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

        TryDifferent.writeExcel(wb, "trying");

    }

    public static String getImgName(String url) {
        String name = url.replace("./test_files/", "");
        name = name.replace(" ", "_");
        return name;
    }

    public static void saveToFile(String in, String out) throws IOException {
        try (InputStream is = new FileInputStream(in); OutputStream os = new FileOutputStream(out)) {
            byte[] b = new byte[2048];
            int length;

            while ((length = is.read(b)) != -1) {
                os.write(b, 0, length);
            }
            System.out.println("saved: " + out);
        }
    }

    public static void saveImgFromUrl(String in, String out) throws IOException {
        URL url = new URL(in);
            try (InputStream is = url.openStream(); OutputStream os = new FileOutputStream(out)) {
                byte[] b = new byte[2048];
                int length;

                while ((length = is.read(b)) != -1) {
                    os.write(b, 0, length);
                }
                System.out.println("saved: " + out);
            }
    }

    private static void print(String msg, Object... args) {
        System.out.println(String.format(msg, args));
    }

    private static String trim(String s, int width) {
        if (s.length() > width)
            return s.substring(0, width-1) + ".";
        else
            return s;
    }



}
