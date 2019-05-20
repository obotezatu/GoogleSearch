package com.alliedtesting.selenium;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.testng.annotations.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GoogleSearch {

	@Test
	public static void main(String[] args) {
		System.setProperty("webdriver.chrome.driver",
				"/media/sef/Archive/ADMINISTRATOR/JAVA/USM_2018/Automation 2019/chromedriver");
		WebDriver driver = new ChromeDriver();
		driver.get("http://www.google.com");
		WebElement element = driver.findElement(By.name("q"));
		element.sendKeys("Cheese!");
		element.submit();
		System.out.println("Page title is: " + driver.getTitle());
		(new WebDriverWait(driver, 10)).until(new ExpectedCondition<Boolean>() {
			public Boolean apply(WebDriver d) {
				return d.getTitle().toLowerCase().startsWith("cheese!");
			}
		});
		List<String> links = getLinksList(driver);
		if (links.size() < 10) {
			WebElement page2 = driver.findElement(By.xpath("//tbody/tr/td[3]"));
			page2.click();
			(new WebDriverWait(driver, 10)).until(new ExpectedCondition<Boolean>() {
				public Boolean apply(WebDriver d) {
					return d.getTitle().toLowerCase().startsWith("cheese!");
				}
			});
			links.addAll(getLinksList(driver));
		}
		writeExcel(analysePage(driver, links));
		driver.quit();
	}

	private static List<String> getLinksList(WebDriver driver) {
		List<WebElement> webElementsLinks = driver.findElements(By
				.xpath("//div[h2[not(contains(text(),'People also ask'))]]//div[@class='r']//a[contains(.,'heese')]"));
		return webElementsLinks.stream().map(el -> el.getAttribute("href")).collect(Collectors.toList());
	}

	private static List<PageInfo> analysePage(WebDriver driver, List<String> links) {
		List<PageInfo> pages = new ArrayList<>();
		for (int i=0; i<10; i++) {
			driver.get(links.get(i));
			(new WebDriverWait(driver, 10)).until(new ExpectedCondition<Boolean>() {
				public Boolean apply(WebDriver d) {
					return d.getTitle().toLowerCase().contains("cheese");
				}
			});
			PageInfo pageInfo = new PageInfo();
			pageInfo.setUrl(driver.getCurrentUrl());
			pageInfo.setTitle(driver.getTitle());
			pageInfo.setOccurrences(StringUtils.countMatches(driver.getPageSource().toLowerCase(), "cheese"));
			pages.add(pageInfo);
		}
		return pages;
	}

	private static void writeExcel(List<PageInfo> pages) {
		try {
			Workbook wb = new XSSFWorkbook();
			Sheet sheet = wb.createSheet("Cheese pages");
			Row rowhead = sheet.createRow(0);
			CellStyle style = wb.createCellStyle();
			Font font = wb.createFont();
			font.setBold(true);
			style.setFont(font);
			rowhead.createCell(0).setCellValue("Title");
			rowhead.createCell(1).setCellValue("URL");
			rowhead.createCell(2).setCellValue("Occurrences of \"cheese\"");

			for (int i = 0; i <= 2; i++) {
				rowhead.getCell(i).setCellStyle(style);
			}
			int rowNumber = 1;
			for (PageInfo page : pages) {
				CellStyle style2 = wb.createCellStyle();
				style2.setWrapText(true);
				Row row = sheet.createRow(rowNumber);
				row.createCell(0).setCellValue(page.getTitle());
				row.createCell(1).setCellValue(page.getUrl());
				row.createCell(2).setCellValue(page.getOccurrences());
				rowNumber++;
			}
			OutputStream fileOut = new FileOutputStream("./src/PagesInfo.xlsx");
			wb.write(fileOut);
			System.out.println("Done.Excel was created");
			wb.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

