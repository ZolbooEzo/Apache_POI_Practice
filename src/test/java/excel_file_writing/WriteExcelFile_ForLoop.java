package excel_file_writing;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile_ForLoop {

	public static void main(String[] args) {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");

		Object[][] empData = { 
				{ "EMP_ID", "Name", "Job" }, 
				{ "101", "Bob", "Engineer" },
				{ "102", "Allen", "Accountant"}, 
				{ "103", "Evy", "Boss" } };

		int rows = empData.length;
		int colums = empData[0].length;


		for (int r = 0; r < rows; r++) {

			XSSFRow row = sheet.createRow(r);

			for (int c = 0; c < colums; c++) {

				XSSFCell cell = row.createCell(c);
				Object value = empData[r][c];

				if (value instanceof String) {
					cell.setCellValue((String) value);
				}
				if (value instanceof Integer) {
					cell.setCellValue((Integer) value);
				}
				if (value instanceof Double) {
					cell.setCellValue((Boolean) value);
				}

			}

		}
		
		try {
		
		String filePath = "src/test/resources/ExcelWriteFolder/empFile.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
		workbook.write(fos);
		fos.close();
		System.out.println("Excel file is written successfully!");
		
		}catch(Exception e) {
			e.printStackTrace();
		}

	}

}
