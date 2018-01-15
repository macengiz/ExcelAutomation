package utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelWritePracice {

	public static void main(String[] args) throws Exception {

		/// .xls older version
		String excelFilePath = "./src/test/resources/TestData/AmazonSearchDataOld2.xls";
		FileInputStream fis = new FileInputStream(excelFilePath);

		// Workbook
		HSSFWorkbook wb1 = new HSSFWorkbook(fis);
		// WorkSheet
		HSSFSheet sh1 = wb1.getSheet("Sheet1");
		HSSFSheet sh2 = wb1.getSheetAt(0);
		// Row

		HSSFRow rw = sh1.getRow(1);
		HSSFCell cell = rw.getCell(1);

		if (cell == null) {
			rw.createCell(1);
			cell.setCellValue("my new value");
		} else {
			cell.setCellValue("my new value");
		}
		FileOutputStream fos = new FileOutputStream(excelFilePath);
		wb1.write(fos);

		fos.close();
		fis.close();
		wb1.close();

	}

}
