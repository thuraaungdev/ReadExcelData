package com.readExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.testng.annotations.Test;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Test
public class ReadExcel {
	/**
	 * @throws IOException
	 */
	public void readExcel() throws IOException {
		String excelPath = "C:\\Users\\User\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx";
		String fileNameString = "TestData";
		String sheetName = "Test";

		// create the Object of file class to get the excel path
		File file = new File(excelPath);

		// To read the file
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum();
		System.out.println("Total rows" + rowCount);
		String data = sheet.getRow(0).getCell(1).getStringCellValue();
		System.out.println(data);

		for (int i = 0; i <= rowCount; i++) {
			Row row = sheet.getRow(i);
			for (int j = 0; j < row.getLastCellNum(); j++) {
				String data1 = sheet.getRow(i).getCell(j).getStringCellValue();
				System.out.println(data1 + "");
			}
			System.out.println();
		}
		wb.close();
	}

}
