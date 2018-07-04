package com.project.Execution;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Properties;
import java.util.Random;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Reporter;
import com.project.util.util;

public class Execute {
	WebDriver driver;
	Properties prop = new Properties();
	InputStream input = null;
	InputStream input1 = null;
	int lastrow;
	int lastcol;
	static HashMap<String, String> hmap = new HashMap<String, String>();
	JavascriptExecutor js = (JavascriptExecutor) driver;
	String excelpath;
	XSSFWorkbook excel;
	WebElement webelement;
	Actions actions;
	public static util utl;

	public void startExecution() throws IOException, InterruptedException, NoSuchMethodException {
		utl = new util();
		excelpath = System.getProperty("user.dir") + "\\input\\" + "inputdata.xlsx";
		excel = utl.setExcelFile(excelpath);
		lastrow = utl.getRowCount("data");

		for (int iCounter = 1; iCounter <= lastrow - 1; iCounter++) {
			if (driver != null) {
				driver.close();
				driver = null;
			}
			lastcol = utl.getColCount("data", 1);
			try {
				loadinputdatarow(iCounter, lastcol);
			} catch (Exception e) {
				System.out.println(e.getMessage());
			}
			OpenUrl(ReadConfig("url"));
			findElement(ReadObjectRepository("lnk_registration")).click();
			selectFromDropdown(ReadObjectRepository("input_searchCountry"), hmap.get("Country"));
			Thread.sleep(2000);
			if (hmap.get("Package").contains("Silver")) {
				findElement(ReadObjectRepository("btn_silver")).click();
			} else {
				findElement(ReadObjectRepository("btn_gold")).click();
			}

			if (hmap.get("Personal#company").contains("Company")) {
				findElement(ReadObjectRepository("lbl_company")).click();
				findElement(ReadObjectRepository("input_companyname")).sendKeys(hmap.get("Companyname"));
				findElement(ReadObjectRepository("input_Businessregnumber")).sendKeys(hmap.get("Businessregnumber"));
			} else {
				findElement(ReadObjectRepository("lbl_personal")).click();
			}
			Thread.sleep(2000);
			findElement(ReadObjectRepository("input_firstName")).sendKeys(hmap.get("Firstname"));
			findElement(ReadObjectRepository("input_lastName")).sendKeys(hmap.get("Lastname"));
			findElement(ReadObjectRepository("alert_acceptcookie")).click();
			findElement(ReadObjectRepository("input_email")).sendKeys(hmap.get("Email"));
			findElement(ReadObjectRepository("input_password")).sendKeys(hmap.get("Password"));
			findElement(ReadObjectRepository("input_confirmPassword")).sendKeys(hmap.get("Confirmpassword"));
			findElement(ReadObjectRepository("input_mobilenumber")).sendKeys(hmap.get("Mobile"));
			findElement(ReadObjectRepository("input_license")).sendKeys(hmap.get("License"));
			scrollobject(ReadObjectRepository("chk_promocode"));
			Thread.sleep(2000);
			// findElement(ReadObjectRepository("btn_register")).click();
			findElement(ReadObjectRepository("lbl_creditcard")).click();
			// Thread.sleep(2000);
			// findElement(ReadObjectRepository("lbl_creditcard")).click();
			if (null == findElement(ReadObjectRepository("btn_register"))) {
				findElement(ReadObjectRepository("lbl_creditcard")).click();
			}
			Thread.sleep(2000);
			scrollobject(ReadObjectRepository("btn_register"));
			Thread.sleep(2000);
			findElement(ReadObjectRepository("btn_register")).click();
			Thread.sleep(2000);
			WebElement emailerr = findElement(ReadObjectRepository("lbl_emailerr"), 1);
			WebElement mobileerr = findElement(ReadObjectRepository("lbl_mobilerr"), 1);
			WebElement invalidmobilerr = findElement(ReadObjectRepository("lbl_invalidmobilerr"), 1);
			WebElement invalidmobillenerr = findElement(ReadObjectRepository("lbl_invalidlengtherr"), 1);

			while (null != mobileerr || null != emailerr || null != invalidmobilerr || null != invalidmobillenerr) {
				if (mobileerr != null) {
					String mobilenumber = hmap.get("Mobile");
					Random r = new Random(System.currentTimeMillis());
					int inum = 1 + r.nextInt(2) * 10000 + r.nextInt(10000);
					mobilenumber = mobilenumber.substring(0, 5) + String.valueOf(inum);
					findElement(ReadObjectRepository("input_mobilenumber")).clear();
					findElement(ReadObjectRepository("input_mobilenumber")).sendKeys(mobilenumber);
				}

				if (invalidmobilerr != null) {
					String mobilenumber = hmap.get("Mobile");
					Random r = new Random(System.currentTimeMillis());
					int inum = 1 + r.nextInt(2) * 10000 + r.nextInt(10000);
					mobilenumber = mobilenumber.substring(0, 5) + String.valueOf(inum);
					findElement(ReadObjectRepository("input_mobilenumber")).clear();
					findElement(ReadObjectRepository("input_mobilenumber")).sendKeys(mobilenumber);
				}

				if (invalidmobillenerr != null) {
					String mobilenumber = hmap.get("Mobile");
					Random r = new Random(System.currentTimeMillis());
					int inum = 1 + r.nextInt(2) * 10000 + r.nextInt(10000);
					mobilenumber = mobilenumber.substring(0, 5) + String.valueOf(inum);
					findElement(ReadObjectRepository("input_mobilenumber")).clear();
					findElement(ReadObjectRepository("input_mobilenumber")).sendKeys(mobilenumber);
				}

				if (emailerr != null) {
					String email = hmap.get("Email");
					Random r = new Random(System.currentTimeMillis());
					int inum = 1 + r.nextInt(2) * 10000 + r.nextInt(10000);
					findElement(ReadObjectRepository("input_email")).clear();
					email = "john.mark" + String.valueOf(inum) + "@gmail.com";
					findElement(ReadObjectRepository("input_email")).sendKeys(email);
				}

				scrollobject(ReadObjectRepository("btn_register"));
				findElement(ReadObjectRepository("btn_register")).click();
				emailerr = findElement(ReadObjectRepository("lbl_emailerr"), 1);
				mobileerr = findElement(ReadObjectRepository("lbl_mobilerr"), 1);
				invalidmobilerr = findElement(ReadObjectRepository("lbl_invalidmobilerr"), 1);
				invalidmobillenerr = findElement(ReadObjectRepository("lbl_invalidlengtherr"), 1);
			}
			// ==================Error
			// Validation================================================
			if (null != findElement(ReadObjectRepository("lbl_err_fieldrequired"), 1)
					|| null != findElement(ReadObjectRepository("lbl_err_password"), 1)
					|| null != findElement(ReadObjectRepository("lbl_err_passworddonotmatch"), 1)
					|| null != findElement(ReadObjectRepository("lbl_err_mobile"), 1)) {

				Reporter.log("Data has not been entered correctly, kindly input valid data");
				Reporter.log("Excution ended");
				// driver.close();

			} else {
				findElement(ReadObjectRepository("lbl_terms")).click();
				Reporter.log("Terms page has been displayed");
				Thread.sleep(2000);
				scrollobject(ReadObjectRepository("btn_agree"));
				findElement(ReadObjectRepository("btn_agree")).click();
				Thread.sleep(2000);
				driver.switchTo().frame("tokenFrame");
				Reporter.log("Credit card page has been displayed");
				findElement(ReadObjectRepository("input_cardnumber")).sendKeys("12345678994561");
				findElement(ReadObjectRepository("input_cardholdername")).sendKeys("park mobile");
				selectValueFromDropdown(ReadObjectRepository("input_expirymonth"), "10");
				selectValueFromDropdown(ReadObjectRepository("input_expiryyear"), "2020");
				findElement(ReadObjectRepository("input_cvc")).sendKeys("123");
				findElement(ReadObjectRepository("btn_submitCreditcard")).click();
				Thread.sleep(5000);
				// driver.close();
			}
		}
	}

	public void scrollobject(String objectname) {
		try {
			webelement = findElement(objectname);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		actions = new Actions(driver);
		actions.moveToElement(webelement);
		actions.perform();
	}

	public static void loadinputdatarow(int currentrow, int lastcol) {
		hmap.clear();
		for (int idata = 0; idata <= lastcol; idata++) {

			/* Adding elements to HashMap */
			try {
				hmap.put(utl.getCellData(0, idata, "data"), utl.getCellData(currentrow, idata, "data"));
			} catch (NoSuchMethodException e) {
				// TODO Auto-generated catch block
			}
		}
	}

	public void OpenUrl(String url) throws IOException {
		try {
			getDriver();
			driver.get(url);
			driver.manage().window().maximize();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public WebDriver getDriver() {
		if (driver == null) {
			System.setProperty("webdriver.chrome.driver",
					System.getProperty("user.dir") + "\\drivers\\" + "chromedriver.exe");
			driver = new ChromeDriver();
		}
		return driver;
	}

	public String ReadObjectRepository(String ObjectName) {
		try {
			if (input1 == null)
				input1 = new FileInputStream("Object_Repository.properties");
			prop.load(input1);
		} catch (IOException e) {
			e.printStackTrace();
		}
		String Object = prop.getProperty(ObjectName);
		return Object;
	}

	public String ReadConfig(String ConfigName) throws IOException {
		try {
			if (input == null) {
				input = new FileInputStream("Config.properties");
				prop.load(input);
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		String Config = prop.getProperty(ConfigName);
		return Config;
	}

	public WebElement findElement(String webelement) throws InterruptedException {
		try {
			WebDriverWait wait = new WebDriverWait(driver, 30);
			WebElement element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(webelement)));
			if (element == null) {
				Reporter.log("Object not Exist ! " + webelement, false);
				Thread.sleep(1000);
				return null;
			}
			Thread.sleep(1000);
			return element;
		} catch (Exception e) {
			Reporter.log("Object not Exist ! " + webelement, false);
			return null;
		}
	}

	public WebElement findElement(String webelement, int waittime) throws InterruptedException {
		WebElement element = null;
		WebDriverWait wait = new WebDriverWait(driver, waittime);

		try {
			element = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(webelement)));
		} catch (Exception e) {
			return element;
		}
		if (element == null) {
			Reporter.log("Object not Exist ! " + webelement, false);
			Thread.sleep(1000);
			return null;
		}
		Thread.sleep(1000);
		return element;
	}

	public void selectValueFromDropdown(String webElement, String value) {
		try {
			Select dropdown = new Select(findElement(webElement));
			dropdown.selectByValue(value);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

	public void selectFromDropdown(String webElement, String value) {
		try {
			WebElement dropdown = findElement(webElement);
			dropdown.click();
			Thread.sleep(1000);
			String xpath = "//div[contains(text(), '" + value + "')]";
			WebElement selectoption = findElement(xpath);
			selectoption.click();
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
