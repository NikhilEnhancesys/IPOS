package com.leadprogram;

//50469

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.SignStyle;
import java.time.temporal.ChronoField;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class LeadProgram extends Launcher {

	public static void downloadMenu() throws InterruptedException, IOException {

		// Menu: Upload Icon click
		driver.findElement(By.xpath("//i[@class='fas fa-upload fa-lg']")).click();

		// Menu: Download Summary Click
		driver.findElement(By.xpath("//a[@href='/upload/main/upload?cat=downloadSummary']")).click();
	}

	public static void leadLeadProgramMenu() throws InterruptedException, IOException {
		// Lead
		driver.findElement(By.xpath(
				"//body/div[@id='main-container']/div[@class='main-container-inner']/div[@id='sidebar']/ul[@class='nav nav-list']/li[6]/a[1]"))
				.click();

		// Lead Program
		driver.findElement(By.xpath("//a[@href='/services/main/service?cat=leadProgram']")).click();
	}

	public static void leadProgram() throws InterruptedException, IOException {

		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
//		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

		// to load the lead program page
		leadLeadProgramMenu();
		Thread.sleep(500);

		// Create btn
		driver.findElement(By.xpath("//button[@id='createLeadProgram']")).click();

		// Lead type
		WebElement ld = driver.findElement(By.xpath("//select[@id='org_type']"));
		Select leadType = new Select(ld);
		leadType.selectByVisibleText("Lead New");

		// Program Name creation
		String pn = "SalesChart";
		// current time HHmm
		ZoneId indianZone = ZoneId.of("Asia/Kolkata");
		// Get the current time in the Indian time zone
		LocalTime currentTime = LocalTime.now(indianZone);
		// Format the time to include only hours and minutes without a colon
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("HHmm");
		String formattedTime = currentTime.format(formatter);
		System.out.println("Indian time: " + formattedTime);
		String programName = pn + formattedTime;
		driver.findElement(By.xpath("//input[@id='prg_name']")).sendKeys(programName);
		System.out.println("Program Name: " + programName);

		// Todays date
		LocalDate today = LocalDate.now();
		// Define the formatter for dd-MM-YYYY
		DateTimeFormatter monthYear = DateTimeFormatter.ofPattern("MMMM yyyy");
		String formattedDate = today.format(monthYear);
		// Print the formatted date
		System.out.println("Today's date in MMMM yyyy format: " + formattedDate);
		// MMMM YYYY
		int Todaysday = today.getDayOfMonth();
		String day = String.valueOf(Todaysday);
		System.out.println(day);

		// Start Date
		try {
			// Locate and click the date input field
			WebElement dateInput = driver.findElement(By.xpath("//input[@id='str_dt']"));
			dateInput.click();
			// Navigate to the correct month/year if needed
			while (!driver.findElement(By.xpath("/html/body/div[9]/div[1]/table/thead/tr[1]/th[2]")).getText()
					.equals("March 2025")) {
				driver.findElement(By.xpath("/html/body/div[9]/div[1]/table/thead/tr[1]/th[3]")).click();
			}
			// Select the specific date (21st)
			List<WebElement> dates = driver.findElements(By.className("day"));
			for (WebElement date : dates) {
				if (date.getText().equals(day)) {
					date.click();
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			System.out.println("Something went wrong while selecting date");
		}

		// End Date
		try {
			// Locate and click the date input field
			WebElement dateInput = driver.findElement(By.xpath("//*[@id=\"end_dt\"]"));
			dateInput.click();
			// Navigate to the correct month/year if needed
			while (!driver.findElement(By.xpath("/html/body/div[9]/div[1]/table/thead/tr[1]/th[2]")).getText()
					.equals("April 2025")) {
				driver.findElement(By.xpath("/html/body/div[9]/div[1]/table/thead/tr[1]/th[3]")).click();
			}
			// Select the specific date (21st)
			List<WebElement> dates = driver.findElements(By.className("day"));
			for (WebElement date : dates) {
				if (date.getText().equals("30")) {
					date.click();
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			System.out.println("Something went wrong while selecting date");
		}
		// Transition Type
		WebElement trtyp = driver.findElement(By.xpath("//select[@id='trans_type']"));
		Select transitionType = new Select(trtyp);
		transitionType.selectByVisibleText("Sub Organization Type");

		// Transition Value
		Actions act = new Actions(driver);
		WebElement trnVal = driver.findElement(By.xpath("//input[@id='trans_value']"));
		trnVal.sendKeys("own");
		act.moveToElement(trnVal).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(1000);
		act.moveToElement(trnVal).sendKeys(Keys.ARROW_DOWN).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(500);

		// Product Knowledge Material
		String filePath = "C:\\Users\\Admin\\Downloads\\sample.pdf";
		WebElement chooseFile = driver.findElement(By.xpath("//input[@id='prod_knwldge_mtrl']"));
		chooseFile.sendKeys(filePath);

		// Call Script
		String filePath2 = "C:\\Users\\Admin\\Downloads\\sample.pdf";
		WebElement chooseFile2 = driver.findElement(By.xpath("//input[@id='call_script']"));
		chooseFile2.sendKeys(filePath2);

		// FAQ
		String filePath3 = "C:\\Users\\Admin\\Downloads\\sample.pdf";
		WebElement chooseFile3 = driver.findElement(By.xpath("//input[@id='faq']"));
		chooseFile3.sendKeys(filePath3);

		// Submit
		driver.findElement(By.xpath("//button[@id='saveLeadProgram']")).click();
		Thread.sleep(500);

		// OK
		driver.findElement(By.xpath("/html/body/div[9]/div/div/div[3]/button[2]")).click();
		Thread.sleep(1000);

		// Back
		driver.findElement(By.xpath("//input[@id='confirmActionLink']")).click();
		Thread.sleep(2000);

		// Searching created program name
		WebElement search = driver.findElement(By.xpath("//input[@id='gs_prg_name']"));
		search.sendKeys(programName);
		Thread.sleep(400);
		act.moveToElement(search).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(400);

		// More
		driver.findElement(By.xpath("//*[@id=\"1\"]/td[13]/div/button")).click();

		// Chaage Status
		driver.findElement(By.xpath("//*[@id=\"1\"]/td[13]/div/ul/li[5]/a")).click();

		// Active
		WebElement sact = driver.findElement(By.xpath("//*[@id=\"beatStatusChanged\"]"));
		Select actve = new Select(sact);
		actve.selectByVisibleText("Active");

		// Comments
		driver.findElement(By.xpath("//*[@id=\"comments\"]")).sendKeys(programName + " " + "Activated");

		// Submit
		driver.findElement(By.xpath("//*[@id=\"changeStatusSubmitButton\"]")).click();

		// Status Change: OK
		Thread.sleep(300);
		driver.findElement(By.xpath("/html/body/div[9]/div/div/div[3]/button[2]")).click();

		// After changing status: Back
		driver.findElement(By.xpath("//*[@id=\"confirmActionLink\"]")).click();

		// Upload Icon -> Download Summary
		downloadMenu();
		Thread.sleep(3000);
		// Refresh
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();

		// Download Form: Download click
		driver.findElement(By.xpath("//button[@id='downloadForm']")).click();

		// Download type
		WebElement downldTypeDrDwn = driver.findElement(By.xpath("//select[@id='select_download_type']"));
		Select XLSXtype = new Select(downldTypeDrDwn);
		XLSXtype.selectByVisibleText("XLSX");

		// Select Template Type
		WebElement selctTempType = driver.findElement(By.xpath("//select[@id='select_template_type']"));
		Select selectTemplateType = new Select(selctTempType);
		selectTemplateType.selectByVisibleText("Prospects Upload");
		Thread.sleep(400);

		// Download Summary: Submit
		driver.findElement(By.xpath("//button[@id='downloadFileSumbit']")).click();

		// Capture the exact time after clicking the button
		LocalDateTime clickTime = LocalDateTime.now();
		DateTimeFormatter formatter3 = DateTimeFormatter.ofPattern("dd-MMM-yyyy HH:mm:ss");
		String formattedClickTime = clickTime.format(formatter3);
		// Print the captured time
		System.out.println("Button clicked at: " + formattedClickTime);

		// Download Summary: Back
		driver.findElement(By.xpath("//input[@id='confirmActionLink']")).click();

		// Refresh
		Thread.sleep(2000);
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();

		// Capturing ProspectUpload_........ file name
		String propspectUploadFileName = driver
				.findElement(By.xpath("//table/tbody/tr[@id='1']//following-sibling::td[2]")).getText();

		// Refresh -> Refresh till status is completed.
		Thread.sleep(3000);
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();

		// Refresh
		Thread.sleep(3000);
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();

		// Download Template: download icon Click
		driver.findElement(By.xpath("//tr[@id='1']//a[@title='Download Template']")).click();

		// Entering values in prospect file
		String propspectUploadFile = "C:\\Users\\Admin\\Downloads\\" + propspectUploadFileName;
		System.out.println("propspectUploadFile: " + propspectUploadFile);
		Thread.sleep(5000);

		String[][] data = { { "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" },
				{ "62435678987654", "ABCX", "JABOTABEK", "NORTH AND CENTRAL JAKARTA", "JAKARTA_JKTNC2" } };

		try {
			// Load the existing Excel file
			FileInputStream fis = new FileInputStream(propspectUploadFile);
			Workbook workbook = new XSSFWorkbook(fis);

			// Access the sheet (replace "Sheet1" with your sheet name)
			Sheet sheet = workbook.getSheet("Sheet1");
			if (sheet == null) {
				sheet = workbook.createSheet("Sheet1");
			}

			// Write the data row by row
			for (int rowIndex = 1; rowIndex <= data.length; rowIndex++) { // Start from the second row (index 1)
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					row = sheet.createRow(rowIndex);
				}

				// Write data into columns
				for (int colIndex = 0; colIndex < data[rowIndex - 1].length; colIndex++) {
					Cell cell = row.getCell(colIndex);
					if (cell == null) {
						cell = row.createCell(colIndex);
					}
					cell.setCellValue(data[rowIndex - 1][colIndex]);
				}
			}

			// Close input stream
			fis.close();

			// Write changes to the file
			FileOutputStream fos = new FileOutputStream(propspectUploadFile);
			workbook.write(fos);
			fos.close();

			System.out.println("Data written successfully to Excel!");

		} catch (IOException e) {
			e.printStackTrace();
		}

		// propspectUploadFile file path after entering data
		System.out.println("propspectUploadFile: " + propspectUploadFile);

		// Upload icon click to upload file
		driver.findElement(By.xpath("//i[@class='fas fa-upload fa-lg']")).click();

		// Upload Summary click to upload excel file
		driver.findElement(By.xpath("//a[@href='/upload/main/upload?cat=uploadSummary']")).click();

		// Refresh to complete the status of uploaded Prospect file
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();
		Thread.sleep(2000);

		// Upload button to upload file
		driver.findElement(By.xpath("//button[@id='uploadForm']")).click();
		Thread.sleep(2000);

		// Select template type to upload file
		WebElement selectTempType = driver.findElement(By.xpath("//select[@id='select_template_type']"));
		Select selectTemplatetype = new Select(selectTempType);
		selectTemplatetype.selectByVisibleText("Prospects Upload");
		Thread.sleep(2000);

		// Choose button web element to upload excel file
		WebElement chooseBtnProspecUpload = driver.findElement(By.xpath("//input[@id='choosen_file_upload']"));
		chooseBtnProspecUpload.sendKeys(propspectUploadFile);
		Thread.sleep(3000);

		// Searching program name to select lead program
		WebElement prgrmSrchToSelectLeadPrgrm = driver.findElement(By.xpath("//input[@id='gs_prg_name']"));
		prgrmSrchToSelectLeadPrgrm.sendKeys(programName);
		act.moveToElement(prgrmSrchToSelectLeadPrgrm).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(500);

		// Radio button to select lead program name
		driver.findElement(By.xpath("//*[@id=\"jqg_list_lead_prg_grid_1\"]")).click();

		// Submit button after uploading excel file
		driver.findElement(By.xpath("//button[@id='uploadFileSumbit']")).click();

		// Back button after uploading Prospect doc
		driver.findElement(By.xpath("//*[@id=\"confirmActionLink\"]")).click();

		// Waiting time to load completely upload summary page
		Thread.sleep(5000);

		// Refresh to complete the status of uploaded Prospect file
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();
		Thread.sleep(2000);

		// Refresh to complete the status of uploaded Prospect file
		driver.findElement(By.xpath("//*[@id=\"refresh_list_commonDiv1\"]")).click();
		Thread.sleep(1000);

		leadLeadProgramMenu();
		Thread.sleep(2000);

		// Searching created program name to search in lead program page
		WebElement searchInLdPrgmPage = driver.findElement(By.xpath("//*[@id=\"gs_prg_name\"]"));
		searchInLdPrgmPage.sendKeys(programName);
		Thread.sleep(500);
		act.moveToElement(searchInLdPrgmPage).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(500);

		// Click on more to Assign Prospect
		driver.findElement(By.xpath("//*[@id=\"1\"]/td[13]/div/button")).click();
		Thread.sleep(500);
		driver.findElement(By.xpath("//*[@id=\"1\"]/td[13]/div/ul/li[3]/a")).click();
		Thread.sleep(3000); // To load the Assign Prospect Grid

		// Activity Date
		try {
			// Locate and click the date input field
			WebElement dateInput = driver.findElement(By.xpath("//*[@id=\"activity_date\"]"));
			dateInput.click();
			// Navigate to the correct month/year if needed
			while (!driver.findElement(By.xpath("/html/body/div[10]/div[1]/table/thead/tr[1]/th[2]")).getText()
					.equals("March 2025")) {
				driver.findElement(By.xpath("/html/body/div[10]/div[1]/table/thead/tr[1]/th[3]")).click();
			}
			// Select the specific date
			List<WebElement> dates = driver.findElements(By.className("day"));
			for (WebElement date : dates) {
				if (date.getText().equals(day)) {
					date.click();
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			System.out.println("Something went wrong while selecting date");
		}

		// Organization selection
		driver.findElement(By.xpath("//*[@id=\"div_selectedDestOrganization\"]/span")).click();
		Thread.sleep(6000); // to load the organizations

		// Entering reference code to select organization
		WebElement refcode = driver.findElement(By.xpath("//*[@id=\"gs_ref_code\"]"));
		refcode.sendKeys("jkpr");
		act.moveToElement(refcode).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(2000);

		// To select organization
		driver.findElement(By.xpath("//*[@id=\"orgSelectionLink\"]")).click();
		Thread.sleep(2000);

		// To select user
		driver.findElement(By.xpath("//*[@id=\"selectedBeatNode\"]")).click();
		Thread.sleep(4000);

		WebElement SA = driver.findElement(By.xpath("//*[@id=\"assignBeatTree_search_input\"]"));
		SA.sendKeys("50469");
		Thread.sleep(500);
		act.moveToElement(SA).sendKeys(Keys.ARROW_DOWN).build().perform();
		act.moveToElement(SA).sendKeys(Keys.RETURN).build().perform();
		Thread.sleep(500);
		// to load the sales agent from the NODE
		Thread.sleep(3000);

		// Radio btn to select sales agent // use this sales agent location(50469)
		driver.findElement(By.xpath("//*[@id=\"userSelElement_1\"]")).click();

		// Select button to select sales agent
		driver.findElement(By.xpath("//*[@id=\"assign-beat-popup\"]/div/div/div[3]/button[2]")).click();

		// check box to select no.of MSISDN
		driver.findElement(By.xpath("//*[@id=\"cb_list_assign_Prospects_grid\"]")).click();

		// Assign button
		driver.findElement(By.xpath("//*[@id=\"submitAssignProspects\"]")).click();
		Thread.sleep(500); // to load the ok button

		// OK to assign confirmation
		driver.findElement(By.xpath("/html/body/div[11]/div/div/div[3]/button[2]")).click();

		// Back after assigning
		driver.findElement(By.xpath("//*[@id=\"confirmActionLink\"]")).click();

		// Search Program name to verify whether LP assigned or not.
		WebElement srchPrgrmNameToVerify = driver.findElement(By.xpath("//*[@id=\"gs_prg_name\"]"));
		srchPrgrmNameToVerify.sendKeys(programName);
		act.moveToElement(srchPrgrmNameToVerify).sendKeys(Keys.RETURN).build().perform();

		// more btn to view prospect
		driver.findElement(By.xpath("//*[@id=\"1\"]/td[13]/div/button")).click();
		Thread.sleep(600);

		// View Prospect
		driver.findElement(By.xpath("//*[@id=\"1\"]/td[13]/div/ul/li[4]/a")).click();

		System.out.println("Lead Program Assigned Successfully✔️");

	}

}
