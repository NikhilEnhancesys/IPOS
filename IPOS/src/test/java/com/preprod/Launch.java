package com.preprod;

import java.time.Duration;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Launch {

	public static WebDriver driver;

	public static void launch() {

		driver = new ChromeDriver();

		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));

		driver.get("https://10.34.57.246:8443/snoc/login?locale=in_ID");
		driver.manage().window().maximize();

		// Advanced btn
		driver.findElement(By.xpath("//*[@id=\"details-button\"]")).click();
		// Proceed link
		driver.findElement(By.xpath("//*[@id=\"proceed-link\"]")).click();

		WebElement lang = driver.findElement(By.xpath("//*[@id=\"languages\"]"));
		Select language = new Select(lang);
		language.selectByVisibleText("English");

		driver.findElement(By.id("username")).sendKeys("EIJKPR");
		driver.findElement(By.id("password")).sendKeys("Admin@123");
		driver.findElement(By.id("submit")).click();
	}

}
