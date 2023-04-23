package writing_excel_file;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile_FormulaCell {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Write_FormulaCell");
		XSSFRow row = sheet.createRow(0);
		
		row.createCell(0).setCellValue(10);
		row.createCell(1).setCellValue(20);
		row.createCell(2).setCellValue(30);
		
		row.createCell(3).setCellFormula("A1+B1+C1");
		
		
		
		try {
			
			String location = "src/test/resources/ExcelWriteFolder/formula_write.xlsx";
			FileOutputStream fos = new FileOutputStream(location);
			
			workbook.write(fos);
			fos.close();
			
			System.out.println("formula_write.xlsx created successfully!");
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		
	}

}
