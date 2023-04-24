package hashmap_excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMap_To_Excel {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Student Data");

		Map<String, String> data = new HashMap<>();

		data.put("101", "John");
		data.put("102", "Smith");
		data.put("103", "Ann");
		data.put("104", "David");
		data.put("105", "Lorde");

		int rowNum = 0;

		for (Map.Entry entry : data.entrySet()) {

			XSSFRow row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue((String) entry.getKey());
			row.createCell(1).setCellValue((String) entry.getValue());
		}

		FileOutputStream fos = new FileOutputStream("src/test/resources/Map_ExcelFolder/MapDataWriting2.xlsx");

		workbook.write(fos);
		fos.close();

		System.out.println("Excel written successful!");

	}

}
