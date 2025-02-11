package com.testcase;

import org.testng.annotations.Test;
import com.readExcel.ExcelLibrary;

public class ReadExcelTest {

    @Test
    public void readExcelTest() throws Exception {
        // Provide the correct Excel file path
        String excelPath = "C:\\Users\\User\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx";
        ExcelLibrary obj = new ExcelLibrary(excelPath);

        // Read data safely
        String dataString = obj.readData("Test", 5, 1);
        System.out.println("The data is: " + dataString);

        // Close workbook after reading
        obj.closeWorkbook();
    }
}
