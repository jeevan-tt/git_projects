package Excel.project;

import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelSheet {

	public static void main(String[] args) throws IOException {

		File excelFile = new File("Contactslist.xlsx");
		FileInputStream fis = new FileInputStream(excelFile);

		// XSSF WK OBJ
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		// sheet

		XSSFSheet sheet = workbook.getSheetAt(0);

		// Iterate rows
		Iterator<Row> rowIt = sheet.iterator();

		while (rowIt.hasNext()) {
			Row row = rowIt.next();

			// iteration of current row
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {

				Cell cell = cellIterator.next();
				System.out.print(cell.toString() + ";");

			}
			System.out.println();
		}

	}

}
