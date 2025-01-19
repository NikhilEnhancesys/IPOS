package com.preprod;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class CustomerInformation extends Launch {

	public static void customerInfo() throws InterruptedException {

		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

		// Nationality
		WebElement SelectNationality = driver.findElement(By.xpath("//*[@id=\"cust_type\"]"));
		Select nationality = new Select(SelectNationality);
		nationality.selectByVisibleText("Indonesian");
		// Read Via Card Reader
		driver.findElement(By.xpath("//*[@id=\"read_Via_Card_Reader\"]")).click();

		// Identity photo
		String filePath = "C:\\Users\\Admin\\Pictures\\Bali.jpg";
		WebElement choose = driver.findElement(By.xpath("//input[@id='nikKtpPhoto_web']")); // conditiom web ele should
																							// be input type // be type
																							// = input
		choose.sendKeys(filePath);

		// Customer Photo
		String filePath2 = "C:\\Users\\Admin\\Pictures\\Customer.jpg";
		WebElement choose2 = driver.findElement(By.xpath("//input[@id='selfiePhoto_web']")); // conditiom web ele should
																								// be type = input
		choose2.sendKeys(filePath2);

		// Customer Name
		driver.findElement(By.xpath("//*[@id=\"customer_name\"]")).sendKeys("Nikhilkumar");
		// Place of Birth
		driver.findElement(By.xpath("//*[@id=\"place_of_birth\"]")).sendKeys("Bali");

		// Date of birth
		driver.findElement(By.xpath("//input[@id='postpaid_dob']")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[11]")));
		String currentMY = driver.findElement(By.xpath("//th[normalize-space()='January 2008']")).getText();
		System.out.println("Outside while: " + currentMY);
		while (true) {
			// Locate the displayed month/year text
			WebElement monthYearLabel = driver.findElement(By.className("datepicker-switch")); // Replace with actual
																								// class
			String displayedMonthYear = monthYearLabel.getText();

			// Break the loop if the desired month and year are displayed
			if (displayedMonthYear.equals("December 2007")) {
				break;
			}

			// Click the "<<" button to navigate backward
			WebElement prevButton = driver.findElement(By.className("prev")); // Replace with actual class
			prevButton.click();
		}
		driver.findElement(By.xpath("//td[@class='day'][normalize-space()='1']")).click();

		// contact No.
		driver.findElement(By.xpath("//*[@id=\"contact_number\"]")).sendKeys("621234567890");

		// Id type
		WebElement IdTypeDrpDwn = driver.findElement(By.xpath("//*[@id=\"id_type\"]"));
		Select IdType = new Select(IdTypeDrpDwn);
		IdType.selectByVisibleText("ID Card (KTP)");
		// Nik and Nokk
		driver.findElement(By.xpath("//*[@id=\"nik_id\"]")).sendKeys("3519101004980002");
		driver.findElement(By.xpath("//*[@id=\"nokk_id\"]")).sendKeys("3519101004980002");
		// Building name and street address
		driver.findElement(By.xpath("//input[@id='building_name_res']")).sendKeys("Enhancesys");
		driver.findElement(By.xpath("//*[@id=\"street_address_res1\"]")).sendKeys("Koramangala");
		// Select Province
		WebElement SelectProvince = driver.findElement(By.xpath("//select[@id='province_res']"));
		Select province = new Select(SelectProvince);
		province.selectByVisibleText("BALI");
		// Select City
		WebElement SelectCity = driver.findElement(By.xpath("//*[@id=\"city_res\"]"));
		Select city = new Select(SelectCity);
		city.selectByVisibleText("BANGLI");
		// Zipcode
		WebElement ZipCodeDrpDwn = driver.findElement(By.xpath("//*[@id=\"postal_code_res\"]"));
		Select zipCode = new Select(ZipCodeDrpDwn);
		zipCode.selectByVisibleText("80611");
		// Reference details
		driver.findElement(By.xpath("//*[@id=\"ref_name\"]")).sendKeys("Anand");
		driver.findElement(By.xpath("//*[@id=\"ref_contact_no\"]")).sendKeys("621234567890");
		// Select Relationship
		WebElement SelectRelationship = driver.findElement(By.xpath("//*[@id=\"ref_relationship\"]"));
		Select relationship = new Select(SelectRelationship);
		relationship.selectByVisibleText("Sibling");
		// Validate btn
		Thread.sleep(500);
		// Verify
		driver.findElement(By.xpath("//button[@id='verifyeKYCId']")).click();
		// Oks
		driver.findElement(By.xpath("/html/body/div[11]/div/div/div[3]/button")).click();
		// Next
		driver.findElement(By.xpath("//*[@id=\"goNextPageId\"]")).click();
		// Download FAB
		driver.findElement(By.xpath("//*[@id=\"download_fab\"]")).click();
		// Download FAB
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"icon\"]")));
		driver.findElement(By.xpath("//*[@id=\"icon\"]")).click();
	}

}
