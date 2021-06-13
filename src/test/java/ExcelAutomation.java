import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelAutomation {

	
	//Identify the TestCases coulmn by scanning entire first row
	//Once column is identified then scan the first column to identify the "Purchase" testcase
	//Once Purchase is identified then fetch all the data of that row and feed into test
	public static void main(String[] args) throws IOException {

		// FileInputStream fis= new
		// FileInputStream("/Volumes/Development/Automation/ExcelDataDriven/ExcelTestData.xlsx");
		// XSSFWorkbook workBook=new XSSFWorkbook(fis);

		XSSFWorkbook workBook = new XSSFWorkbook("/Volumes/Development/Automation/ExcelDataDriven/ExcelTestData.xlsx");

		int sheetCount = workBook.getNumberOfSheets();

		for (int i = 0; i < sheetCount; i++) {
			if (workBook.getSheetAt(i).getSheetName().equalsIgnoreCase("TestData")) {
				XSSFSheet sheet = workBook.getSheetAt(i);
				
			Iterator<Row>row=	sheet.iterator();
				
			Row firstRow=row.next();
			
		Iterator <Cell> cell=	firstRow.cellIterator();
			
while(cell.hasNext())
{
	Cell value=cell.next();
	if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
	{
		//	
	}
}
			
			
			}
		}

	}

}
