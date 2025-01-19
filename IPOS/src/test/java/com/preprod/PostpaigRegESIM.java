package com.preprod;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class PostpaigRegESIM extends Launch {

	public static void postpaidRegistrationESIM() throws InterruptedException {

		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

		// Postpaid tile
		wait.until(ExpectedConditions
				.visibilityOfElementLocated(By.xpath("//*[@id=\"dasboardServiceIconsDetailsDiv\"]/div[7]")));
		driver.findElement(By.xpath("//*[@id=\"dasboardServiceIconsDetailsDiv\"]/div[7]")).click();
		// Stock location
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//select[@id='stock_Location']")));
		WebElement StkLocDrpDwn = driver.findElement(By.xpath("//select[@id='stock_Location']"));
		Select OwnGeraiStock = new Select(StkLocDrpDwn);
		OwnGeraiStock.selectByVisibleText("Own Gerai Stock");
		// MSISDN Search
		driver.findElement(By.xpath("//*[@id=\"msisdnSearchButton\"]")).click();
		// Select MSISDN
		driver.findElement(By.xpath("//*[@id=\"jqg_list_bill_payment_list_1\"]")).click();
		// Proceed
		driver.findElement(By.xpath("//*[@id=\"modal-SearchMsisdn\"]/div/div/div[2]/div/div/div/div[2]/div/button"))
				.click();
		// ESIM
		driver.findElement(By.xpath("//*[@id=\"ESIM\"]")).click();

		Actions act = new Actions(driver);
		// Post paid Package
		WebElement EsimPp = driver.findElement(By.xpath("//*[@id=\"2\"]/td[3]/label/input"));
		EsimPp.sendKeys("Prime 70 GB-Contract 12");
		// enter key after searching product
		act.moveToElement(EsimPp).sendKeys(Keys.RETURN).perform();
		Thread.sleep(1000);
		// arrow down key and enter key for selecting listed products
		act.moveToElement(EsimPp).sendKeys(Keys.ARROW_DOWN).sendKeys(Keys.RETURN).build().perform();

		Thread.sleep(1000);
		// SIM Card
		WebElement SIMCard = driver.findElement(By.xpath("//*[@id=\"3\"]/td[3]/label/input"));
		SIMCard.sendKeys("ESIM");
		act.moveToElement(SIMCard).sendKeys(Keys.RETURN).perform();
		Thread.sleep(500);
		act.moveToElement(SIMCard).sendKeys(Keys.ARROW_DOWN).sendKeys(Keys.RETURN).build().perform();
		// NEXT btn
		driver.findElement(By.xpath("//*[@id=\"goNextPageId\"]")).click();
		// Email
		driver.findElement(By.xpath("//*[@id=\"EmailId\"]")).sendKeys("no@no.com");
		// Send OTP btn
		driver.findElement(By.xpath("//*[@id=\"sendOtpBtn\"]")).click();
		// Enter OTP
		driver.findElement(By.xpath("//*[@id=\"OTPValue\"]")).sendKeys("1234");
		// Validate btn
		driver.findElement(By.xpath("//*[@id=\"validateOTP\"]/button")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[11]/div/div/div[3]/button")));
		// OK
		driver.findElement(By.xpath("/html/body/div[11]/div/div/div[3]/button")).click();
	}

	public static void main(String[] args) throws InterruptedException {
		Launch.launch();
		postpaidRegistrationESIM();
		CustomerInformation.customerInfo();
//		System.out.println("ESIM Postpaid Registration without deposit done successfully");

	}

}
