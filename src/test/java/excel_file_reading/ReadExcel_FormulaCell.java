package excel_file_reading;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel_FormulaCell {

	public static void main(String[] args) {

		try {
			FileInputStream fis = new FileInputStream(Constants.formulaExcelFile);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet("Sheet1");

			int rowNum = sheet.getLastRowNum();
			int colNum = sheet.getRow(0).getLastCellNum();
			
			
			for (int r = 0; r <= rowNum; r++) {
				XSSFRow row = sheet.getRow(r);
				
				for(int c = 0; c < colNum; c++) {
					
					XSSFCell cell = row.getCell(c);
					
					switch(cell.getCellType()) {
					
					case STRING: System.out.print(cell.getStringCellValue()); break;
					case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
					case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
					case FORMULA: System.out.print(cell.getNumericCellValue()); break;
					}
					System.out.print(" | ");
				}
				System.out.println();
				
			}
			
			fis.close();
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
