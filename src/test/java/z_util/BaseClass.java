package z_util;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	
	private static WebDriver driver;
	
	@BeforeMethod
	public static WebDriver getDriver() {
		
		if(driver == null) {
			
			String browser = "chrome";
			switch(browser) {
			
			case "chrome": 
				WebDriverManager.chromedriver().setup();
				driver = new ChromeDriver();
				break;
				
			case "firefox":
				WebDriverManager.firefoxdriver().setup();
				driver = new FirefoxDriver();
				break;
			}
			
			driver.get("https://practice.automationtesting.in/my-account/");
			driver.manage().window().maximize();
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			
		}
		
		return driver;
	}
	
	@AfterMethod
	public static void closeDriver() {
		
		if(driver != null) {
			driver.quit();
			driver = null;
		}
		
	}
	

}
