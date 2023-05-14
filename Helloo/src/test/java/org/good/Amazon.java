package org.good;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Amazon {
	public static void main(String[] args) throws IOException {
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.amazon.in/");

		WebElement txtsearchBox = driver.findElement(By.id("twotabsearchtextbox"));
		txtsearchBox.sendKeys("iphone", Keys.ENTER);

		List<WebElement> allphonenames = driver
				.findElements(By.xpath("//span[@class='a-size-medium a-color-base a-text-normal']"));

		List<WebElement> allphoneprices = driver.findElements(By.xpath("//span[@class='a-price-whole']"));

		File file = new File("C:\\Users\\HP\\eclipse-workspace\\Helloo\\exceldata\\createsheet.xlsx");

		FileInputStream stream = new FileInputStream(file);

		Workbook book = new XSSFWorkbook(stream);
		Sheet createSheet = book.createSheet("iphonenames 1");

		for (int i = 0; i < allphonenames.size(); i++) {
			WebElement element = allphonenames.get(i);
			String text = element.getText();
			System.out.println(text);
			Row createRow = createSheet.createRow(i);
			Cell createCell = createRow.createCell(0);
			createCell.setCellValue(text);
			WebElement element2 = allphoneprices.get(i);
			String text2 = element2.getText();
			System.out.println(text2);
			Row createRow2 = createSheet.createRow(i);
			Cell createCell2 = createRow2.createCell(1);
			createCell2.setCellValue(text2);

		}

		FileOutputStream stream2 = new FileOutputStream(file);
		book.write(stream2);
		System.out.println("Done ok!!!");
	}

}
