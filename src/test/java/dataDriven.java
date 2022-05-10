import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {

	public ArrayList<String> getData(String testCaseName) throws IOException {
		// Identify testcases column by scanning the entire 1st row
		// Once column is identified then scan entire testcase column to identify
		// purchase testcase row
		// Pull all the data of purchase row and feed into test

		// we have to pass file path as below concept:
		// fileInputStream Argument has access to read any file

		// fileInputStream argument
		ArrayList<String> a = new ArrayList<String>();

		FileInputStream fis = new FileInputStream(
				"C:\\Users\\Swapn\\Desktop\\QA Material Review\\Rahul Shetty\\Selenium\\ExcelDataDrivenTestingUtilities\\ExcelDriven.xlsx");

		// workbook have ability to read excel file
		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		// get number of sheet in excel
		int sheets = workbook.getNumberOfSheets();
		for (int i = 0; i < sheets; i++) {
			// get sheet via index from sheets
			if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				// identify testcase column by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();// sheet collection of rows
				Row firstrow = rows.next();
				Iterator<Cell> cell = firstrow.cellIterator(); // collection of cells
				int k = 0;
				int column = 0;
				while (cell.hasNext()) {
					Cell value = cell.next();
					if (value.getStringCellValue().equalsIgnoreCase(testCaseName)) // mentioning cell data here to find
																					// column
					{
						column = k;
					}
					k++;
				}
				System.out.println(column);

				// Once column is identified then scan entire testcase column to identify
				while (rows.hasNext()) {
					Row r = rows.next(); // on first row control
					if (r.getCell(column).getStringCellValue().equalsIgnoreCase(testCaseName)) {
						// after you grab testcase row then pull all data of that row and feed into test
						Iterator<Cell> cv = r.cellIterator();
						while (cv.hasNext()) {

							Cell c = cv.next();
							if (c.getCellType() == CellType.STRING) {
								a.add(c.getStringCellValue()); // Add content to array
							} else {
								// convert to string
								String b = NumberToTextConverter.toText(c.getNumericCellValue());
								// Add to arraylist
								a.add(b);
							}
						}
					}

				}
			}
		}
		return a;
	}
}
