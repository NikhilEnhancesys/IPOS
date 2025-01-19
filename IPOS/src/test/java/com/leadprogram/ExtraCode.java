package com.leadprogram;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;

public class ExtraCode {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		
//		// Open the Excel file
//		FileInputStream fileInputStream = new FileInputStream(new File(propspectUploadFile));
//		Workbook workbook = new XSSFWorkbook(propspectUploadFile);
//
//		// Get the Sheet (Sheet1)
//		Sheet sheet = workbook.getSheet("Sheet1");
//
//		// Data to be written (example)
//		String[][] data = { { "MSISDN", "CUSTOMER_NAME", "REGION", "AREA", "CLUSTER" }, // Header row (optional, if you
//																						// want to overwrite it)
//				{ "621234567890", "Adhi", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" }, // Row 2
//				{ "621234567890", "Anwar", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" }, // Row 3
//				{ "621234567890", "Ahmad", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" } // Row 4
//		};
//		// Write data starting from the 2nd row (index 1)
//		int startingRow = 1; // 0-based indexing, so row 2 means index 1
//		for (int i = 0; i < data.length; i++) {
//			Row row = sheet.getRow(startingRow + i);
//			if (row == null) {
//				row = sheet.createRow(startingRow + i); // Create row if it doesn't exist
//			}
//
//			for (int j = 0; j < data[i].length; j++) {
//				Cell cell = row.getCell(j);
//				if (cell == null) {
//					cell = row.createCell(j); // Create cell if it doesn't exist
//				}
//				cell.setCellValue(data[i][j]); // Set the cell value
//			}
//		}
//		// Close the input stream
//		fileInputStream.close();
//		// Save the changes to the Excel file
//		FileOutputStream fileOutputStream = new FileOutputStream(new File(propspectUploadFile));
//		workbook.write(fileOutputStream);
//		// Close resources
//		fileOutputStream.close();
//		workbook.close();
//		System.out.println("Data written successfully to the Excel file.");
		
		// ------------------------------------------------------------------------
		
//		// Searching created program name to search in lead program page
//				WebElement searchInLdPrgmPagetoEnsure = driver.findElement(By.xpath("//*[@id=\"gs_prg_name\"]"));
//				searchInLdPrgmPagetoEnsure.sendKeys(programName);
//				Thread.sleep(500);
//				act.moveToElement(searchInLdPrgmPagetoEnsure).sendKeys(Keys.RETURN).build().perform();
//				Thread.sleep(500);

	}

}
