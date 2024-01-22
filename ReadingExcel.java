package automation;

import java.io.File;
import java.io.FileInputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadingExcel {

	public static void main(String[] args) throws Exception {
		// Specify the location of excel file
		File src = new File("D:\\Training Materials//Short Notes.xlsx");

		//load file
		FileInputStream fis = new FileInputStream(src);

		//load Workbook
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		//Load worksheet
		XSSFSheet sh = wb.getSheet("Sheet1");

		//print the loaded sheet name
		System.out.println(sh.getSheetName());

		//print selenium from excel sheet
		System.out.println(sh.getRow(1).getCell(1).getStringCellValue());

		//print java from ecxel sheet
		System.out.println(sh.getRow(2).getCell(1).getStringCellValue());

		//print float/double from excel sheet
		System.out.println(sh.getRow(1).getCell(3).getNumericCellValue());


		//get integer from excel sheet
		//System.out.println(int) sh.getRow(1).getCell(4).getNumericCellValue());

		//total no. of row
		System.out.println("Total Rows : " + sh.getPhysicalNumberOfRows());

		//print total no. of row 2nd way
		System.out.println("Total Rows : " + (sh.getLastRowNum() + 1));

		//print total number of colum 1st way
		System.out.println("Total Columns : " + sh.getRow(1).getPhysicalNumberOfCells());

		//print total number of columns 2nd way
		System.out.println("Total columns : " + sh.getRow(1).getLastCellNum());


		//real time implementation
		System.setProperty("webdriver.chrome.driver", "D:\\Training Materials\\workspace\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		driver.get("https://www.facebook.com/");

     // Enter username using excel file
		String abc = sh.getRow(2).getCell(3).getStringCellValue();
		driver.findElement(By.id("email")).sendKeys(abc);

	}

}									
