import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelAutomation {

	// Identify the TestCases coulmn by scanning entire first row
	// Once column is identified then scan the first column to identify the
	// "Purchase" testcase
	// Once Purchase is identified then fetch all the data of that row and feed into
	// test

	public static void main(String[] args) throws IOException {

		ExcelAutomation excel = new ExcelAutomation();
		ArrayList<String> al = excel.getData("Purchase");
		System.out.println(al);

	}

	public ArrayList<String> getData(String testCaseName) throws IOException {
		// FileInputStream fis= new
		// FileInputStream("/Volumes/Development/Automation/ExcelDataDriven/ExcelTestData.xlsx");
		// XSSFWorkbook workBook=new XSSFWorkbook(fis);

		int index = 0;
		XSSFWorkbook workBook = new XSSFWorkbook("/Volumes/Development/Automation/ExcelDataDriven/ExcelTestData.xlsx");
		ArrayList<String> al = new ArrayList<String>();

		int sheetCount = workBook.getNumberOfSheets();

		for (int i = 0; i < sheetCount; i++) {
			if (workBook.getSheetAt(i).getSheetName().equalsIgnoreCase("TestData")) {
				XSSFSheet sheet = workBook.getSheetAt(i);

				Iterator<Row> row = sheet.iterator();

				Row firstRow = row.next();

				Iterator<Cell> cell = firstRow.cellIterator();

				while (cell.hasNext()) {
					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						index = value.getColumnIndex();
					}

				}

				while (row.hasNext()) {
					Row r = row.next();
					String rowValue = r.getCell(index).getStringCellValue();
					if (rowValue.equalsIgnoreCase(testCaseName)) {
						Iterator<Cell> c = r.cellIterator();
						while (c.hasNext()) {
							Cell c1 = c.next();
							al.add(c1.getStringCellValue());

						}
					}

				}

			}
		}
		return al;
	}

}
