package writing_excel_file;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile_ForEachLoop {

	public static void main(String[] args) {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");

		Object[][] empData = { 
				{ "EMP_ID", "Name", "Job" }, 
				{ 101, "Bob", "SDET" },
				{ 102, "Allen", "SDE"}, 
				{ 103, "Evy", "Master" } };

		int rowCount = 0;
		
		for(Object emp[] : empData) {
			
			XSSFRow row = sheet.createRow(rowCount++);
			
			int columnCount = 0;
			
				for(Object value : emp) {
					
					XSSFCell cell = row.createCell(columnCount++);
					
					if(value instanceof String) {
						cell.setCellValue((String) value);
					}
					if(value instanceof Integer) {
						cell.setCellValue((Integer) value);
					}
					if(value instanceof Boolean) {
						cell.setCellValue((Boolean) value);
					}
				}
		}
		
		try {
		
		String filePath = "src/test/resources/ExcelWriteFolder/empFileForEach.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
		workbook.write(fos);
		fos.close();
		System.out.println("Excel file is written successfully!");
		
		}catch(Exception e) {
			e.printStackTrace();
		}

	}

}
