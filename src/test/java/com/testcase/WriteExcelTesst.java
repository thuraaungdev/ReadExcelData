package com.testcase;

import org.testng.annotations.Test;

import com.writeExcel.WriteExcel;

public class WriteExcelTesst {
	WriteExcel obj= new WriteExcel();
	@Test
	public void writeExcelTest() throws Exception {
		obj.writeExcel("Test","Male", 0, 2);
	}
	@Test
	public void writeExcelTest1() throws Exception {
		obj.writeExcel("Test","Female", 1, 2);
	}
}
