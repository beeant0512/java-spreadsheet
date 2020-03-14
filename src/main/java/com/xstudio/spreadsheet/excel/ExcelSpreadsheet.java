package com.xstudio.spreadsheet.excel;

import com.xstudio.spreadsheet.excel.entity.Header;
import com.xstudio.spreadsheet.excel.entity.Row;
import com.xstudio.spreadsheet.excel.entity.WorkSheet;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

/**
 * @author xiaobiao
 * @version 2020/3/14
 */
public class ExcelSpreadsheet {
    private static final String NEW_LINE = System.getProperty("line.separator");

    public static void write(File file, WorkSheet sheet) throws IOException {
        // creates a FileWriter Object
        FileWriter writer = new FileWriter(file);
        // write spreed sheet header
        // workbook start tag
        writeXmlHeader(writer);

        // write spread sheet
        writeSheet(writer, sheet);
        // write spread sheet footer
        // workbook close tag
        writeXmlFooter(writer);

        writer.close();
    }

    private static void writeSheet(FileWriter writer, WorkSheet sheet) throws IOException {
        List<Header> headers = sheet.getHeaders();
        // worksheet name start tag
        writer.write(indentation("<Worksheet ss:Name=\"" + sheet.getName() + "\">", 0));

        // table start tag
        writer.write(indentation("<Table>", 1));

        // write headers
        StringBuilder columnStr = new StringBuilder();
        StringBuilder headerCell = new StringBuilder(indentation("<Row>", 2));

        int idx = 1;
        for (Header header : headers) {
            columnStr.append(column(header, idx));
            headerCell.append(cell(header, header.getValue()));
            idx += 1;
        }

        headerCell.append(indentation("</Row>", 2));
        writer.write(columnStr.toString());
        writer.write(headerCell.toString());

        // write value
        List<Row> rows = sheet.getRows();
        for (Row row : rows) {
            writer.write(indentation("<Row>", 2));
            for (Header header : headers) {
                writer.write(cell(header, row.getCells().get(header.getKey())));
            }
            writer.write(indentation("</Row>", 2));
        }


        // table close tag
        writer.write(indentation("</Table>", 1));

        // worksheet name close tag
        writer.write(indentation("</Worksheet>", 0));
    }

    private static String column(Header header, int idx) {
        StringBuilder sb = new StringBuilder("<Column ss:Index=\"");
        sb.append(idx);
        sb.append("\" ss:AutoFitWidth=\"0\" ss:Width=\"");
        sb.append(header.getWidth());
        sb.append("\" />");

        return indentation(sb.toString(), 2);
    }

    private static String cell(Header header, Object value) {
        StringBuilder sb = new StringBuilder("<Cell><Data ss:Type=\"");
        sb.append(header.getType());
        sb.append("\">");
        sb.append(value);
        sb.append("</Data></Cell>");
        return indentation(sb.toString(), 3);
    }

    private static void writeXmlHeader(FileWriter writer) throws IOException {
        writer.write(indentation("<?xml version=\"1.0\" encoding=\"UTF-8\"?>", 0));
        writer.write(indentation("<?mso-application progid=\"Excel.Sheet\"?>", 0));
        writer.write(indentation("<Workbook", 0));
        writer.write(indentation(" xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"", 0));
        writer.write(indentation(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"", 0));
        writer.write(indentation(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"", 0));
        writer.write(indentation(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"", 0));
        writer.write(indentation(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">", 0));
        writer.flush();
    }

    private static void writeXmlFooter(FileWriter writer) throws IOException {
        writer.write("</Workbook>");
        writer.flush();
    }

    private static String indentation(String value, int size) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < size; i++) {
            sb.append("    ");
        }

        sb.append(value);
        sb.append(NEW_LINE);
        return sb.toString();
    }
}
