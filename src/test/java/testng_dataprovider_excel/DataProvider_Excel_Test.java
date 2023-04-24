package testng_dataprovider_excel;

import java.io.IOException;

import org.openqa.selenium.By;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import z_util.BaseClass;
import z_util.XLUtility;

public class DataProvider_Excel_Test extends BaseClass{

	@Test(dataProvider = "usernameAndPass")
	public void logintest(String username, String pass) throws InterruptedException, IOException {
		
		BaseClass.getDriver().findElement(By.xpath("//input[@id='username']")).sendKeys(username);
		BaseClass.getDriver().findElement(By.xpath("//input[@id='password']")).sendKeys(pass);		
		Thread.sleep(1000);
		
	}
	
	
	@DataProvider(name = "usernameAndPass")
	public String[][] getUserNameAndPasswordData() throws IOException{
		
		String path = "src/test/resources/ExcelFolder/username_password.xlsx";
		XLUtility xlutil = new XLUtility(path);
		
		int totalRows = xlutil.getRowCount("Sheet1");
		int totalCols = xlutil.getCellCount("Sheet1", 1);
		
		String loginData[][] = new String[totalRows][totalCols];
		
		for(int r = 1; r <= totalRows; r++) {
			for(int c = 0; c < totalCols; c++) {
				loginData[r-1][c] = xlutil.getCellData("Sheet1", r, c);
			}
		}
		return loginData;
	}
	
}
