package com.readExcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelLibrary {
    private XSSFWorkbook wb;
    private XSSFSheet sheet;

    // Constructor to initialize workbook
    public ExcelLibrary(String excelPath) throws IOException {
        File file = new File(excelPath);
        FileInputStream fis = new FileInputStream(file);
        wb = new XSSFWorkbook(fis);
        fis.close(); // Close input stream
    }

    // Method to read data from Excel safely
    public String readData(String sheetName, int row, int col) {
        sheet = wb.getSheet(sheetName);
        if (sheet == null) {
            return "Error: Sheet '" + sheetName + "' not found!";
        }

        Row rowData = sheet.getRow(row);
        if (rowData == null) {
            return "Error: Row " + row + " not found!";
        }

        Cell cell = rowData.getCell(col);
        if (cell == null) {
            return "Error: Cell (" + row + "," + col + ") is empty!";
        }

        // Use DataFormatter to handle different cell types
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    // Close workbook
    public void closeWorkbook() throws IOException {
        if (wb != null) {
            wb.close();
        }
    }
}
