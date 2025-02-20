package Sector_Details;

import org.testng.annotations.AfterMethod;
import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;
import java.time.Duration;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.AfterClass;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class Sector_Class {
	public WebDriver driver;
	
	  @BeforeMethod
	@BeforeClass
	    public void setUp() {
	        // Set the path to the chromedriver executable
	        driver = new ChromeDriver();
	        
	        // Maximize the browser
	        driver.manage().window().maximize();
	    }
	  @SuppressWarnings("deprecation")
	@Test
	  public void openMyBlog() throws InterruptedException {
		
		//launching application-wagix funds
		System.out.println("Launching browser...");
		driver.get("https://www.franklintempleton.com/investments/options/mutual-funds/products/90040/IS/western-asset-income-fund/WAGIX");
		//wait to load the page
		System.out.println("wating..");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		//accept the cookies
		System.out.println("accepting..");
		WebElement OkButton=driver.findElement(By.xpath("//*[@id=\"onetrust-accept-btn-handler\"]"));
		OkButton.click();
		//wait to load
		System.out.println("waiting..");
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		//WebElement Heading =  driver.findElement(By.xpath("//h1[text()=' WAGIX ']"));
		//click on portfolio option
		System.out.println("clicking on portfolio..");
		WebElement Portfolio =driver.findElement(By.xpath("//a/span[text()='Portfolio']"));
		Portfolio.click();
		
		//scrolling down the page till the table
		
		   
		 
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));

		System.out.println("clicking on sector");
		
		WebElement sector = driver.findElement(By.xpath("//a[text()=' Sector ']"));
		Actions actions = new Actions(driver);
        actions.moveToElement(sector).perform();

        // Perform actions on the element
        sector.click();
		
		int no_of_rows=driver.findElements(By.xpath("((//table[@class='table table--secondary tableExpand bar-chart-comp ng-star-inserted']/tbody)[2]/tr/td[2]/span/span[2])")).size();		
		
		//saving data
		System.out.println("writing data...");
		Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sectors Data");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("Sectors");
        row.createCell(1).setCellValue("Funds");
        
		for(int i=1;i<=no_of_rows; i++) {
			Row rows =sheet.createRow(i);
			String Sector_Name = driver.findElement(By.xpath("((//table[@class='table table--secondary tableExpand bar-chart-comp ng-star-inserted']/tbody)[2]/tr/td[2]/span/span[2])["+ i +"]")).getText();
			System.out.println("Row..:"+ Sector_Name);
			rows.createCell(0).setCellValue(Sector_Name);
        	String fund= driver.findElement(By.xpath("((//table[@class='table table--secondary tableExpand bar-chart-comp ng-star-inserted']/tbody)[2]/tr/td[3])["+i+"]")).getText();
			System.out.println("cell..:"+ fund);
        	rows.createCell(1).setCellValue(fund);
        		
		}
        try (FileOutputStream fileOut = new FileOutputStream("SectorsData.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Closing the workbook..");
     // Close the workbook and WebDriver
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
		
	  }
		
		  @AfterMethod
		@AfterClass public void tearDown() 
		  { 
			  driver.close(); 
			  }
		 
}
