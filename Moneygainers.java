package api;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Moneygainers {

	public static void main(String[] args) throws InterruptedException, IOException {

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\vino\\Documents\\chromedriver.exe");
//        ChromeOptions options = new ChromeOptions();
//
//        options.addArguments("headless");
//
//        options.addArguments("window-size=1200x600");

		ChromeDriver driver = new ChromeDriver();
		driver.get("https://money.rediff.com/gainers");
		Thread.sleep(3000);
		// driver.manage().window().maximize();
		String title = driver.getTitle();

		System.out.println(title);

		driver.manage().timeouts().implicitlyWait(05, TimeUnit.SECONDS);

		File src = new File("C:\\Users\\vino\\Desktop\\Testdata\\mgNew.xlsx");
		FileInputStream fis = new FileInputStream(src);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// XSSFSheet sheet = workbook.createSheet("MoneyGainersList");
		XSSFSheet sheet = workbook.getSheet("MoneyGainersList");
		if(sheet == null)
	    sheet = workbook.createSheet("MoneyGainersList");

		XSSFFont font = workbook.createFont();
		XSSFCellStyle style = workbook.createCellStyle();

		Row row;
		Cell cell;

		// *[@id="leftcontainer"]/table/tbody/tr[4]

		String bfrcomp_nmae = "//*[@id=\"leftcontainer\"]/table/tbody/tr[";
		String aftrcomp_name = "]/td/a";

		String bfrcomp_grp = "//*[@id=\"leftcontainer\"]/table/tbody/tr[";
		String aftrcomp_grp = "]/td[2]";

		String bfrcomp_prevpr = "//*[@id=\"leftcontainer\"]/table/tbody/tr[";
		String aftrcomp_prevpr = "]/td[3]";

		String bfrcomp_currpr = "//*[@id=\"leftcontainer\"]/table/tbody/tr[";
		String aftrcomp_currpr = "]/td[4]";

		String bfrcomp_gain = "//*[@id=\"leftcontainer\"]/table/tbody/tr[";
		String aftrcomp_gain = "]/td[5]";

		int rownum = 0;

		for (int i = 1; i <= 8; i++) {
			row = sheet.createRow(rownum++);
			String actual_cmp = bfrcomp_nmae + i + aftrcomp_name;
			String compnsme = driver.findElement(By.xpath(actual_cmp)).getText();
			cell = row.createCell(0);
			
			if (cell != null) {				
				    cell = row.createCell(0);				
				   }				
				   cell.setCellValue(driver.findElement(By.xpath(actual_cmp)).getText());
			

			String actual_grp = bfrcomp_grp + i + aftrcomp_grp;
			String group = driver.findElement(By.xpath(actual_grp)).getText();
			cell = row.createCell(1);
			cell.setCellValue(group);

			String actual_prevpr = bfrcomp_prevpr + i + aftrcomp_prevpr;
			// String prev_price = driver.findElement(By.xpath(actual_prevpr)).getText();
			cell = row.createCell(2);
			cell.setCellValue(driver.findElement(By.xpath(actual_prevpr)).getText());

			String actual_currpr = bfrcomp_currpr + i + aftrcomp_currpr;
			String curr_price = driver.findElement(By.xpath(actual_currpr)).getText();
			cell = row.createCell(3);
			row.setRowStyle(style);
			cell.setCellValue(curr_price);

			String actual_gain = bfrcomp_gain + i + aftrcomp_gain;
			String gain = driver.findElement(By.xpath(actual_gain)).getText();
			cell = row.createCell(4);
			cell.setCellValue(gain);

			System.out.println( compnsme +"<------>"+curr_price +"<------>"+ gain);

		}

		FileOutputStream fileOut = new FileOutputStream(src);
		workbook.write(fileOut);
		fileOut.close();
		driver.close();
		driver.quit();

	}
}
//		
//		
//		
//		 XSSFRow row = sheet.getRow(1);	
//		 XSSFCell cell = row.getCell(0);				
//		   if (cell == null) {				
//		    cell = row.createCell(0);				
//		   }				
//		   cell.setCellValue(CompanyName);
//		
//	   // System.out.println(rowSize);
//	    
////	   for (int i =0; i<=5;i++) {
////		   for (WebElement alllinks :allRows ) {
////				  String link = alllinks.getText();
////				  System.out.println(link);
//		   
//	   
//	    
//	  for (WebElement alllinks :allRows ) {
//		  String link = alllinks.getText();
//		  System.out.println(link);
//	 
//	  }

//		
//		Row row = sheet.createRow(0);
//		// Create a cell and put a value in it.
//		Cell cell = row.createCell(0);
//		cell.setCellValue();

//List<WebElement> elemen = driver.findElementsByXPath("//*[@id=\"leftcontainer\"]/table/tbody/tr[10]/preceding-sibling::*");
//for (WebElement CompanyName : elemen) {
//	
//	System.out.println(CompanyName.getText());
//
//	row = sheet.createRow(rownum++);
//	cell = row.createCell(0);
//	cell.setCellValue((Date) driver.findElementsByXPath("//*[@id=\"leftcontainer\"]/table/tbody/tr[10]/preceding-sibling::*"));

//		
//		FileOutputStream fileOut = new FileOutputStream(src);
//		workbook.write(fileOut);
//		fileOut.close();

//	
//	}
