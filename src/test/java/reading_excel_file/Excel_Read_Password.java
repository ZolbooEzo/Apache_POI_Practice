package reading_excel_file;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Read_Password {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		String loopType = "Loop";

		FileInputStream fis = new FileInputStream(Constants.passwordFile);
		String password = "12345";

		XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis, password);
		XSSFSheet sheet = workbook.getSheetAt(0);

		int rows = sheet.getLastRowNum();
		int colNum = sheet.getRow(0).getLastCellNum();

		if (loopType == "Loop") {

			// ---------------------------------------------------------------------------------------------------------

			for (int r = 0; r <= rows; r++) {
				XSSFRow row = sheet.getRow(r);

				for (int c = 0; c < colNum; c++) {
					XSSFCell cell = row.getCell(c);

					switch (cell.getCellType()) {

					case STRING: System.out.print(cell.getStringCellValue()); break;
					case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
					case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
					case FORMULA: System.out.print(cell.getNumericCellValue()); break;

					}
					System.out.print(" | ");
				}
				System.out.println();
			}
			workbook.close();

			System.out.println("For loop executed");

		} else if (loopType == "Iterator") {

			// ---------------------------------------------------------------------------------------------------------

			Iterator<Row> iterator = sheet.iterator();

			while (iterator.hasNext()) {

				Row nextRow = iterator.next();
				Iterator<Cell> cellIterator = nextRow.cellIterator();

				while (cellIterator.hasNext()) {

					Cell cell = cellIterator.next();
					switch (cell.getCellType()) {

					case STRING: System.out.print(cell.getStringCellValue()); break;
					case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
					case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
					case FORMULA: System.out.print(cell.getNumericCellValue()); break;

					}
					System.out.print(" | ");
				}
				System.out.println();
			}
			workbook.close();

			System.out.println("Iterator executed");

		} else {
			System.out.println("Please choose *Loop* or *Iterator* on line number 20");
		}

	}

}
