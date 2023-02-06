package Excel.project;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadSheetPoi {

	public static void main(String[] args) throws IOException {

		String exceLPath = "C:\\Users\\jeeva\\OneDrive\\Desktop\\jeevan.xlsx";
		FileInputStream fileInputStream = new FileInputStream(exceLPath);
		String password = "";

		XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
		XSSFSheet sheet = null;

		sheet = workbook.getSheet("Contacts");

		if (null != sheet && sheet.getProtect()) {

			if (sheet.validateSheetPassword(password)) {

				int rows = sheet.getLastRowNum();
				int cols = sheet.getRow(1).getLastCellNum();
				for (int r = 0; r < rows; r++) {
					XSSFRow row = sheet.getRow(r);
					for (int c = 0; c < cols; c++) {
						XSSFCell cell = row.getCell(c);

						System.out.print(cell.getStringCellValue() + "\t");

					}
					System.out.println();

				}
			} else {
				System.out.println("File Mismatch");
			}

		} else {
			System.out.println("UN AUTHERISED");
		}
	}

}
