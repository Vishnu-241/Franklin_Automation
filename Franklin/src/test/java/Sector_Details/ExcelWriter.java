package Sector_Details;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;


public class ExcelWriter {
    public static void main(String[] args) {
        // Create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Portfolio Data");

        // Create a row and put some cells in it
        Row row = sheet.createRow(0);
        Cell cell1 = row.createCell(0);
        cell1.setCellValue("Test Case");
        Cell cell2 = row.createCell(1);
        cell2.setCellValue("Result");

        // Write some test data
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("Login Test");
        row1.createCell(1).setCellValue("Pass");

        // Write the output to a file
        try (FileOutputStream fileOut = new FileOutputStream("TestData.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

