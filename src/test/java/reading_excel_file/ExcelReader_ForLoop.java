package reading_excel_file;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader_ForLoop {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream(Constants.excelFilePath);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		int rows = sheet.getLastRowNum();
		int colums = sheet.getRow(1).getLastCellNum();

		for (int r = 0; r <= rows; r++) {

			XSSFRow row = sheet.getRow(r);
			for (int c = 0; c < colums; c++) {

				XSSFCell cell = row.getCell(c);

				switch (cell.getCellType()) {

				case STRING:
					System.out.print(cell.getStringCellValue() + " | ");
					break;

				case NUMERIC:
					System.out.print(cell.getNumericCellValue() + " | ");
					break;

				case BOOLEAN:
					System.out.print(cell.getBooleanCellValue() + " | ");
					break;
				}

			}
			
			System.out.println();

		}

	}

}
