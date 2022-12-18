package Excel.project;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	private static String[] columns = { "First Name", "Last Name", "Email", "Date Of Birth" };
	private static List<Contacts> contacts = new ArrayList<Contacts>();

	public static void main(String[] args) throws IOException {
		contacts.add(new Contacts("Jeevan", "AT", "jeevantt4@gmail.com", "03-06-19999"));
		contacts.add(new Contacts("aeevan", "BT", "Aeevantt4@gmail.com", "03-06-19998"));
		contacts.add(new Contacts("beevan", "CT", "Beevantt4@gmail.com", "03-06-19996"));
		contacts.add(new Contacts("ceevan", "DT", "Ceevantt4@gmail.com", "03-06-19996"));
		contacts.add(new Contacts("deevan", "ET", "Deevantt4@gmail.com", "03-06-19997"));
		contacts.add(new Contacts("eeevan", "TT", "Eeevantt4@gmail.com", "03-06-19995"));
		contacts.add(new Contacts("feevan", "GT", "Feevantt4@gmail.com", "03-06-19994"));
		contacts.add(new Contacts("geevan", "JT", "Geevantt4@gmail.com", "03-06-19993"));
		contacts.add(new Contacts("heevan", "LT", "Heevantt4@gmail.com", "03-06-19992"));
		contacts.add(new Contacts("ieevan", "TT", "Ieevantt4@gmail.com", "03-06-19991"));
		contacts.add(new Contacts("keevan", "ST", "leevantt4@gmail.com", "03-06-19990"));
		contacts.add(new Contacts("leevan", "TT", "aeevantt4@gmail.com", "03-06-19993"));
		contacts.add(new Contacts("meevan", "XT", "geevantt4@gmail.com", "03-06-19995"));
		contacts.add(new Contacts("neevan", "LT", "teevantt4@gmail.com", "03-06-19996"));

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Contacts");

		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.RED.getIndex());

		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Create a Row
		Row headerRow = sheet.createRow(0);

		for (int i = 0; i < columns.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(columns[i]);
			cell.setCellStyle(headerCellStyle);

		}

		// Create Other rows and cells with contacts data
		int rowNum = 1;

		for (Contacts contact : contacts) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(contact.firstName);

			row.createCell(1).setCellValue(contact.lastName);
			row.createCell(2).setCellValue(contact.email);
			row.createCell(3).setCellValue(contact.dateOfBirth);
		}

		// Resize all columns to fit the content size
		for (int i = 0; i < columns.length; i++) {
			sheet.autoSizeColumn(i);

		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("Downloadscontacts.xlsx");
		workbook.write(fileOut);
		fileOut.close();
	}

}