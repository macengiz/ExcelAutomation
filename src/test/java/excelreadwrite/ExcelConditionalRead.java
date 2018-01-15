package excelreadwrite;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelConditionalRead {

	public static void main(String[] args) throws IOException {

		String excelPath = "C:/Users/aydin/Desktop/EmpData.xlsx";

		FileInputStream in = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new XSSFWorkbook(in);
		XSSFSheet worksheet = workbook.getSheet("Sheet2");

		int rowsCount = worksheet.getPhysicalNumberOfRows();

		for (int rownum = 1; rownum < rowsCount; rownum++) {
			String execute = worksheet.getRow(rownum).getCell(0).toString();
			if (execute.equals("Y")) {
				String searchItem = worksheet.getRow(rownum).getCell(1).toString();
				System.out.println("searching for " + searchItem);
			}

		}

		in.close();
		workbook.close();

	}

}
