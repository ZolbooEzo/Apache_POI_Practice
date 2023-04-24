package excel_file_writing;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel_file_reading.Constants;

public class WriteExcelFile_ApplyFormula {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream(Constants.applyFormula);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		
		sheet.getRow(6).getCell(0).setCellFormula("SUM(A2:A5)");
		fis.close();
		
		FileOutputStream fos = new FileOutputStream(Constants.applyFormula);
		workbook.write(fos);
		workbook.close();
		fos.close();
		
		System.out.println("Update write sucessful!");
		
		
	}

}
