package com.writeExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {
	public void writeExcel(String sheetName, String cellValue, int row, int col) throws Exception {
		String excelPath="C:\\Users\\User\\eclipse-workspace\\ReadExcel\\TestData\\TestData.xlsx";
		File file =new File (excelPath);
		  FileInputStream fis = new FileInputStream(file);
		 XSSFWorkbook wb = new XSSFWorkbook(fis);
	     XSSFSheet sheet= wb.getSheet(sheetName);
	     sheet.getRow(row).createCell(col).setCellValue(cellValue);
	     FileOutputStream fos= new FileOutputStream(new File(excelPath));
	     
	     wb.write(fos);
	     wb.close();
	     fos.close();
	}

}
