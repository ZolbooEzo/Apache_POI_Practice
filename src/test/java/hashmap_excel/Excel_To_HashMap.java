package hashmap_excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel_file_reading.Constants;

public class Excel_To_HashMap {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream(Constants.mapExcelData);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Student Data");
		
		int rowNum = sheet.getLastRowNum();
		
		
		HashMap<String, String> data = new HashMap<>();
		
		for(int r = 0; r <= rowNum; r++) {
			String key = sheet.getRow(r).getCell(0).getStringCellValue();
			String value = sheet.getRow(r).getCell(1).getStringCellValue();
			data.put(key, value);
		}
		
		
		// This is just printing hash map Key and value side by side
		
		for(Map.Entry entry : data.entrySet()) {
			
			System.out.println(entry.getKey() + " " + entry.getValue());
			
		}
		
		
		
		
		

	}

}
