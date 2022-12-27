package Excel.project;

import java.io.FileInputStream;
import java.io.IOException;

import javax.print.DocFlavor.STRING;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
public class ReadSheetPoi {

	public static void main(String[] args) throws IOException {

		String exceLPath = ".\\data\\Contactslist.xlsx";
		FileInputStream fileInputStream = new FileInputStream(exceLPath);

		
		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = workbook.getSheet("Cobtacts");
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(1).getLastCellNum();
		for (int r = 0; r <= rows; r++) {
			XSSFRow row = sheet.getRow(r);
			for (int c = 0; c <= cols; c++) {
				XSSFCell cell = row.getCell(c);
				cell.getBooleanCellValue();
				
				System.out.println();

			}

		}

	}

}
