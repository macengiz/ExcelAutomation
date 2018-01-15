package excelreadwrite;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {

		String excelPath = "C:/Users/aydin/Desktop/EmpData.xlsx";

		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet2");

		int rowsCount = worksheet.getPhysicalNumberOfRows();

		System.out.println("rows count: " + rowsCount);

		XSSFCell cell = worksheet.getRow(1).getCell(2);

		if (cell == null) {
			cell = worksheet.getRow(1).createCell(2);
		}
		cell.setCellValue("Fail");

		cell = worksheet.getRow(5).getCell(2);

		if (cell == null) {
			cell = worksheet.getRow(5).createCell(2);
		}
		cell.setCellValue("Pass");

		FileOutputStream out = new FileOutputStream(excelPath);
		workbook.write(out);

		out.close();
		in.close();
		workbook.close();

	}
}
