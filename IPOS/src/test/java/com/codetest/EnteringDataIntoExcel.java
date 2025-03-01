//package com.codetest;
//
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//
//public class EnteringDataIntoExcel {
//
//	public static void main(String[] args) {
//
//		String filePath = "C:\\Users\\Admin\\Downloads\\ProspectsUpload_26012025_213055.xlsx"; // Path to your Excel
//																								// file
//
//		String[][] data = { { "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
//				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
//				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
//				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
//				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" } };
//
//		try {
//			// Load the existing Excel file
//			FileInputStream fis = new FileInputStream(filePath);
//			Workbook workbook = new XSSFWorkbook(fis);
//
//			// Access the sheet (replace "Sheet1" with your sheet name)
//			Sheet sheet = workbook.getSheet("Sheet1");
//			if (sheet == null) {
//				sheet = workbook.createSheet("Sheet1");
//			}
//
//			// Write the data row by row
//			for (int rowIndex = 1; rowIndex <= data.length; rowIndex++) { // Start from the second row (index 1)
//				Row row = sheet.getRow(rowIndex);
//				if (row == null) {
//					row = sheet.createRow(rowIndex);
//				}
//
//				// Write data into columns
//				for (int colIndex = 0; colIndex < data[rowIndex - 1].length; colIndex++) {
//					Cell cell = row.getCell(colIndex);
//					if (cell == null) {
//						cell = row.createCell(colIndex);
//					}
//					cell.setCellValue(data[rowIndex - 1][colIndex]);
//				}
//			}
//
//			// Close input stream
//			fis.close();
//
//			// Write changes to the file
//			FileOutputStream fos = new FileOutputStream(filePath);
//			workbook.write(fos);
//			fos.close();
//
//			System.out.println("Data written successfully to Excel!");
//
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
//	}
//}
