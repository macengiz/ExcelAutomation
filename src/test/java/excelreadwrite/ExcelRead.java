package excelreadwrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {

		String excelPath = "C:/Users/aydin/Desktop/EmpData.xlsx";

		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet1");

		int rowsCount = worksheet.getPhysicalNumberOfRows();
		System.out.println("number of rows " + rowsCount);
		System.out.println(worksheet.getRow(0).getCell(0));
		System.out.println(worksheet.getRow(1).getCell(0));
		System.out.println(worksheet.getRow(3).getCell(1));

		String cellValue = worksheet.getRow(3).getCell(1).toString();
		System.out.println(cellValue);

		int sheetrowsCount = worksheet.getPhysicalNumberOfRows();
		for (int row = 1; row < sheetrowsCount; row++) {
			String name = worksheet.getRow(row).getCell(1).toString();
			String department = worksheet.getRow(row).getCell(2).getStringCellValue();
			String id = worksheet.getRow(row).getCell(0).toString();
			System.out.println(name);
			System.out.println(id + "--> " + name + "--> " + department);
		}

		in.close();
		workbook.close();

	}

}
