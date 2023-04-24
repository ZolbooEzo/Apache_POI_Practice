package testng_dataprovider_excel;

import org.openqa.selenium.By;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import z_util.BaseClass;

public class DataProvider_HardCode_Test extends BaseClass{
	
	@Test(dataProvider = "usernameAndPass")
	public void logintest(String username, String pass) throws InterruptedException {
		
		BaseClass.getDriver().findElement(By.xpath("//input[@id='username']")).sendKeys(username);
		BaseClass.getDriver().findElement(By.xpath("//input[@id='password']")).sendKeys(pass);		
		
		Thread.sleep(2000);
	}
	
	
	@DataProvider(name = "usernameAndPass")
	public String[][] getUserNameAndPasswordData() {
		String loginData[][] = {
				{"pass", "test"},
				{"code", "test"},
				{"keys", "test"}
		};
		return loginData;
	}
	
	
	
	

}
