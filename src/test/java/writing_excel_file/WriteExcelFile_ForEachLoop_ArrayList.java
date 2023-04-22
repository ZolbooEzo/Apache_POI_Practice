package writing_excel_file;

import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelFile_ForEachLoop_ArrayList {

	public static void main(String[] args) {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");

		ArrayList <Object[]> empData = new ArrayList<Object[]>();
		
		empData.add(new Object[] {"EmpID", "Name", "Job"});
		empData.add(new Object[] {201, "Evan", "Technician"});
		empData.add(new Object[] {202, "Ivar", "Doctor"});
		empData.add(new Object[] {203, "David", "Streamer"});

		int rowCount = 0;
		
		for(Object[] emp : empData) {
			
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
		
		String filePath = "src/test/resources/ExcelWriteFolder/empFileForEach_ArrayList.xlsx";
		FileOutputStream fos = new FileOutputStream(filePath);
		workbook.write(fos);
		fos.close();
		System.out.println("Excel file is written successfully!");
		
		}catch(Exception e) {
			e.printStackTrace();
		}

	}

}
