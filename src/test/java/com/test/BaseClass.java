package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriver.Navigation;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	
		ChromeDriver driver;
		
		//1
		public void getDriver() {
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
		}//2
		public void loadurl(String url) {
			driver.get(url);
		}//3
		public void maximize() {
			driver.manage().window().maximize();
		}//4
		public void minimize() {
			driver.manage().window().minimize();
		}//5
		public Alert stichWindowToAlert() {
			Alert alert = driver.switchTo().alert();
			return alert;
		}//6
		public Alert clickOkInAlert() {
			Alert alert = driver.switchTo().alert();
			alert.accept();
			return alert;
		}//7
		public Alert clickCancelInAlert() {
			Alert alert = driver.switchTo().alert();
			alert.dismiss();
			return alert;
		}//8
		public Alert enterTextInAlert(String text) {
			Alert alert = driver.switchTo().alert();
			alert.sendKeys(text);
			return alert;
		}//9
		public void type(WebElement element,String data) {
		element.sendKeys(data);
		}//10
		public void closeAllWindow() {
			driver.quit();
			}//11
		public void closecurrentWindow() {
			driver.close();
			}//12
		public void click(WebElement element) {
			element.click();
		}//13
		public String getTitle() {
			String title=driver.getTitle();
			return title;
		}//14
		public String geturl() {
			String currentUrl=driver.getCurrentUrl();
			return currentUrl;
		}//15
		public String getText(WebElement element) {
			String text=element.getText();
			return text;
		}//16
		public String getAttribute(WebElement element) {
			String attribute=element.getAttribute("value");
			return attribute;
		}//17
		public String getAttribute(WebElement element, String attributeValue) {
			String attribute=element.getAttribute(attributeValue);
			return attribute;
		}//18
		public WebElement findElementById(String attributeValue) {
			WebElement element = driver.findElement(By.id(attributeValue));
			return element;
		}//19
		public WebElement findElementByname(String attributeValue) {
			WebElement element = driver.findElement(By.name(attributeValue));
			return element;
		}//20
		public WebElement findElementByClassName(String attributeValue) {
			WebElement element = driver.findElement(By.className(attributeValue));
			return element;
		}//21
		public WebElement findElementByXpath(String xPath) {
			WebElement element = driver.findElement(By.xpath(xPath));
			return element;
		}//22
		public void selectOptionByIndex(WebElement element,int cardType) {
			Select select=new Select(element);
			select.selectByIndex(cardType);
			}//23
		public void selectOptionByValue(WebElement element,String attributrvalue) {
			Select select=new Select(element);
			select.selectByValue(attributrvalue);
			}//24
		public void selectOptionByText(WebElement element,String text) {
			Select select=new Select(element);
			select.selectByVisibleText(text);
			}//25
		public void deSelectOptionByIndex(WebElement element,int index) {
			Select select=new Select(element);
			select.deselectByIndex(index);
		}//26
		public void deSelectOptionByValue(WebElement element,String attributrvalue) {
			Select select=new Select(element);
			select.deselectByValue(attributrvalue);
			}//27
		public void deselectOptionByText(WebElement element,String text) {
			Select select=new Select(element);
			select.deselectByVisibleText(text);
			}//28
		public void deSelectAll(WebElement element) {
			Select select=new Select(element);
			select.deselectAll();
			}//29
		public void typeJs(WebElement element,String data) {
			JavascriptExecutor executor=(JavascriptExecutor)driver;
			executor.executeScript("arguments[0],setattribute('value',"+data+"'", element);
			}//30
		public void clickUsingJs(WebElement element) {
			JavascriptExecutor executor=(JavascriptExecutor)driver;
			executor.executeScript("arguments[0].click()", element);
			}//31
		public WebDriver switchToFrameByIndex(int index) {
			WebDriver frame = driver.switchTo().frame(index);
			return frame;
		}//32
		public WebDriver switchToFrameByFrameId(String id) {
			WebDriver frame = driver.switchTo().frame(id);
			return frame;
		}//33
		//ScreenShot    
		public File screenShot(String dest) {
		 TakesScreenshot screenShot=(TakesScreenshot)driver;
		 File flie = screenShot.getScreenshotAs(OutputType.FILE);
		return flie;
		}//34
		//ScreenShot of particular element by id
		public WebElement screenShotCurrentElement(String id) {
			WebElement element = findElementById(id);
			 File flie = element.getScreenshotAs(OutputType.FILE);
			return element;
			}//35
		//ScreenShot of particular element by ClassName
		public WebElement screenShotElementByClassname(String className) {
			WebElement element = findElementByClassName(className);
			File flie = element.getScreenshotAs(OutputType.FILE);
			return element;
		}//36
		//ScreenShot of particular element by ClassName
			public WebElement screenShotElementByname(String Name) {
				WebElement element = findElementByClassName(Name);
				File flie = element.getScreenshotAs(OutputType.FILE);
				return element;
			}//37
			//ScreenShot of particular element by xpath
			public WebElement call(String xpath) {
				WebElement element = findElementByXpath(xpath);
				File flie = element.getScreenshotAs(OutputType.FILE);
				return element;
			}	//38	
		//navigateto
			public Navigation navigateTo(String url) {
			Navigation navigate = driver.navigate();
			navigate.to(url);
			return navigate;
			}//39
			//refresh
			public Navigation refresh() {
				Navigation navigate = driver.navigate();
				navigate.refresh();
				return navigate;
				}
		//
			
			public String getdata(String sheetname,int rownum,int cellnum) throws IOException {
			String data = null;
			File file = new File("C:\\Users\\sadhana\\eclipse-workspace\\Maven\\Excel\\excellTask.xlsx");
			FileInputStream stream = new FileInputStream(file);
			Workbook workbook = new XSSFWorkbook(stream);
			Sheet sheet = workbook.getSheet(sheetname);
			Row row = sheet.getRow(rownum);
			Cell cell = row.getCell(cellnum);
			CellType cellType = cell.getCellType();
			
			switch (cellType) {
			case STRING:
				data = cell.getStringCellValue();
				break;
			case NUMERIC:
				if(DateUtil.isCellInternalDateFormatted(cell)) {
					java.util.Date dateCellValue = cell.getDateCellValue();
					SimpleDateFormat dateformat = new SimpleDateFormat("dd-MMM-yy");
					data = dateformat.format(dateCellValue);
				}
				else {
					double d = cell.getNumericCellValue();
					BigDecimal b = BigDecimal.valueOf(d);
					data = b.toString();
				}
				break;
	default:
		break;
			}
			return data;
		}

		
	public void writeNewCell(String sheet,int rono, int cellNO,String data) throws IOException {
		
		
		File file=new File("C:\\Users\\sadhana\\eclipse-workspace\\Maven\\Excel\\excellTask.xlsx");
		FileInputStream stream=new FileInputStream(file);
		Workbook book=new XSSFWorkbook(stream);
		Sheet sheetName = book.getSheet(sheet);
		Row row = sheetName.getRow(rono);
		Cell cell = row.createCell(cellNO);
		cell.setCellValue(data);
		
		FileOutputStream o=new FileOutputStream(file);
		book.write(o);
	}
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		



}
