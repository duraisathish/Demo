package read;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class Demo {

	String driverPath = "C:\\";
	public WebDriver driver;

	@BeforeClass
	public void setup() {
		System.out.println("*****************");
		System.out.println("Launching Chrome Broswer");
		System.setProperty("webdriver.chrome.driver", driverPath + "chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	public void readExcel(String filePath, String fileName, String sheetName) throws IOException, InterruptedException {

		// Create an object of File class to open xlsx file

		File file = new File("D:\\ExportExcel.xlsx");

		// Create an object of FileInputStream class to read excel file

		FileInputStream inputStream = new FileInputStream(file);

		Workbook guru99Workbook = null;

		// Find the file extension by splitting file name in substring and getting only
		// extension name

		String fileExtensionName = fileName.substring(fileName.indexOf("."));

		// Check condition if the file is xlsx file

		if (fileExtensionName.equals(".xlsx")) {

			// If it is xlsx file then create object of XSSFWorkbook class

			guru99Workbook = new XSSFWorkbook(inputStream);

		}

		// Check condition if the file is xls file

		else if (fileExtensionName.equals(".xls")) {

			// If it is xls file then create object of XSSFWorkbook class

			guru99Workbook = new HSSFWorkbook(inputStream);

		}

		// Read sheet inside the workbook by its name

		Sheet guru99Sheet = guru99Workbook.getSheet(sheetName);

		for (Row row : guru99Sheet) {
			for (Cell cell : row) {
				DataFormatter fmt = new DataFormatter();
				String cellValue = fmt.formatCellValue(cell);
				if (cellValue != null && !cellValue.isEmpty()) {

					setup();

					Thread.sleep(10000);

					driver.get(cellValue);

					Thread.sleep(10000);

					Login();

					Thread.sleep(20000);

					ClickOpportunityOwner();

					Thread.sleep(10000);

					AMUserDetails();

					Thread.sleep(10000);

					AMUserLogin();

					Thread.sleep(20000);

					driver.get(cellValue);

					Thread.sleep(25000);

					Analysis();

					Thread.sleep(25000);

					ClickAnalysis();

					Thread.sleep(20000);

					SaveAnalyis();

					Thread.sleep(10000);

					driver.close();

					Set<String> st = driver.getWindowHandles();
					Iterator<String> it = st.iterator();
					String parent = it.next();
					driver.switchTo().window(parent);

					driver.close();

					System.out.println("Completed");
					System.out.print(cellValue + "\t");
				}

			}
			System.out.println();
		}
	}

	// Main function is calling readExcel function to read data from excel file

	public static void main(String... strings) throws IOException, InterruptedException {

		// Create an object of Demo class

		Demo objExcelFile = new Demo();

		// Prepare the path of excel file

		String filePath = System.getProperty("user.dir") + "\\src\\excelExportAndFileIO";

		// Call read file method of the class to read data

		objExcelFile.readExcel(filePath, "ExportExcel.xlsx", "Demo");

	}

	@Test(priority = 1, enabled = true)
	public void Login() {
		driver.manage().window().maximize();

		driver.findElement(By.id("username")).sendKeys("sajiv@minusculetechnologies.com.qa");

		driver.findElement(By.id("password")).sendKeys("alkabila@0408");

		driver.findElement(By.id("Login")).click();
	}

	@Test(priority = 2, enabled = true)
	public void ClickOpportunityOwner() {

		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}

		driver.switchTo().frame(0);

		WebDriverWait wait3 = new WebDriverWait(driver, 60);
		wait3.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"ownerNameSpan\"]/a")));

		WebElement owner = driver.findElement(By.xpath("//*[@id=\"ownerNameSpan\"]/a"));
		owner.click();

	}

	@Test(priority = 3, enabled = true)
	public void AMUserDetails() throws InterruptedException {
		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}

		Thread.sleep(10000);

		Set<String> st = driver.getWindowHandles();
		Iterator<String> it = st.iterator();
		String parent = it.next();
		String child = it.next();
		driver.switchTo().window(parent);
		driver.switchTo().window(child);

		WebDriverWait wait5 = new WebDriverWait(driver, 60);
		wait5.until(ExpectedConditions.elementToBeClickable(By.partialLinkText("User Detail")));

		WebElement userdetails = driver.findElement(By.partialLinkText("User Detail"));
		userdetails.click();
	}

	@Test(priority = 4, enabled = true)
	public void AMUserLogin() {
		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}

		try {
			driver.switchTo().frame(0);

			WebDriverWait wait5 = new WebDriverWait(driver, 60);
			wait5.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"topButtonRow\"]/input[4]")));

			WebElement login = driver.findElement(By.xpath("//*[@id=\"topButtonRow\"]/input[4]"));
			login.click();
		} catch (Exception e) {

		}
	}

	@Test(priority = 5, enabled = true)
	public void Analysis() {

		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}

		((JavascriptExecutor) driver).executeScript("scroll(0,100)");

		try {
			driver.switchTo().frame(0);

			driver.findElement(By.xpath("//*[@id=\"opportunityApp\"]/div[1]/div[5]/ul/li[2]/a"));

			WebDriverWait wait3 = new WebDriverWait(driver, 60);
			wait3.until(
					ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='opportunityApp']/div[1]/div[5]/ul/li")));

			List<WebElement> element1 = driver.findElements(By.xpath("//*[@id='opportunityApp']/div[1]/div[5]/ul/li"));
			for (WebElement elements : element1) {

				WebElement test = elements.findElement(By.tagName("a"));

				if (test.getText().contains("ANALYSIS")) {
					test.click();
				}

			}
		} catch (Exception e) {
		}

	}

	@Test(priority = 6, enabled = true)
	public void ClickAnalysis() {
		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}
		((JavascriptExecutor) driver).executeScript("scroll(0,100)");

		try {

			driver.findElement(By.xpath("//*[@id=\"opportunityApp\"]/div[1]/div[5]/ul/li[2]/a"));

			WebDriverWait wait3 = new WebDriverWait(driver, 60);
			wait3.until(
					ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='opportunityApp']/div[1]/div[5]/ul/li")));

			List<WebElement> element1 = driver.findElements(By.xpath("//*[@id='opportunityApp']/div[1]/div[5]/ul/li"));
			for (WebElement elements : element1) {

				WebElement test = elements.findElement(By.tagName("a"));

				if (test.getText().contains("ANALYSIS")) {
					test.click();
				}

			}
		} catch (Exception e) {
		}
	}

	@Test(priority = 7, enabled = true)
	public void SaveAnalyis() {
		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}
		try {
			driver.switchTo().frame(driver.findElement(By.id("iframeLeaseAnalysis")));

			WebElement Save = driver.findElement(
					By.xpath("//*[@id=\"leaseAnalysisApp\"]/ng-include/div/div/div[1]/div[2]/div/button[3]/span"));
			Save.click();
		} catch (Exception e) {
		}
	}

	@Test(priority = 8, enabled = true)
	public void UserLogout() {
		try {
			Thread.sleep(2000);
		} catch (Exception e) {
			System.out.println("Thread Error");
		}

		try {
			driver.switchTo().defaultContent();

			WebDriverWait wait4 = new WebDriverWait(driver, 60);
			wait4.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"oneHeader\"]/div[1]/div/a")));

			WebElement logout = driver.findElement(By.xpath("//*[@id=\"oneHeader\"]/div[1]/div/a"));
			logout.click();
		} catch (Exception e) {
		}

	}

}
