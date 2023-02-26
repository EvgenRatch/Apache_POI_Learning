package com.example.apache_poi_learning.excelOperations;

import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.Iterator;

public class ReadingExcel {
    @SneakyThrows
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\chang\\Desktop\\countries.xlsx";
        FileInputStream fileInputStream = new FileInputStream(excelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet1 = workbook.getSheetAt(0); // or use getSheetAt(index);
        int rows = sheet1.getLastRowNum();
        int columns = sheet1.getRow(1).getLastCellNum();

//        for (int r = 0; r < rows; r++) {
//            XSSFRow row = sheet1.getRow(r);
//            for (int c = 0; c < columns; c++) {
//                XSSFCell cell = row.getCell(c);
//                switch (cell.getCellType()) {
//                    case STRING -> System.out.print(cell.getStringCellValue());
//                    case NUMERIC -> System.out.print(cell.getNumericCellValue());
//                    case BOOLEAN -> System.out.print(cell.getBooleanCellValue());
//                }
//                System.out.print(" | ");
//            }
//            System.out.println();
//        }
        for (Row currentRow : sheet1) {
            XSSFRow row = (XSSFRow) currentRow;
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();
                switch (cell.getCellType()) {
                    case STRING -> System.out.print(cell.getStringCellValue());
                    case NUMERIC -> System.out.print(cell.getNumericCellValue());
                    case BOOLEAN -> System.out.print(cell.getBooleanCellValue());
                }
                System.out.print(" | ");
            }
            System.out.println();
        }

    }
}
