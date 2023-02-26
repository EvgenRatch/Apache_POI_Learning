package com.example.apache_poi_learning.excelOperations;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

@Slf4j
public class WritingExcel {

    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet1 = workbook.createSheet("Employee information");
        String filePath = "C:\\Users\\chang\\Desktop\\employee2.xlsx";
        Object[][] employeeData = {
                {
                        "EmployeeID", "Name", "Job", "Salary"
                },
                {
                        "1", "David", "Developer", 2315
                },
                {
                        "2", "Maxim", "Developer", 4315
                },
                {
                        "3", "Eugene", "Developer", 4315
                }
        };
        int rows = employeeData.length;
        int cells = employeeData[0].length;
        for (int r = 0; r < rows; r++) {
            XSSFRow row = sheet1.createRow(r);
            for (int c = 0; c < cells; c++) {
                XSSFCell cell = row.createCell(c);
                var value = employeeData[r][c];
                if (value instanceof String) {
                    cell.setCellValue((String) value);
                }
                if (value instanceof Integer) {
                    cell.setCellValue((Integer) (value));
                }
                if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) (value));
                }
            }
        }
        try (OutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        } catch (IOException exception) {
            log.info("An exception occurred : " + exception);
        }
    }
}
