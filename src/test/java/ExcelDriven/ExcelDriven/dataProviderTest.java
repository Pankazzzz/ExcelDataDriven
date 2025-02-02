package ExcelDriven.ExcelDriven;

import java.io.File;
import java.io.FileInputStream;
import java.util.Formatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class dataProviderTest {
	
	DataFormatter formatter = new DataFormatter();
	
	@Test(dataProvider ="driveTest")
	public void testData(String name, String hobby, String id,String job)
	{
		System.out.println(name+" "+hobby+" "+id+" "+job);
	}
	
	@org.testng.annotations.DataProvider(name="driveTest")
	public Object[][] getData() throws Exception
	{
		FileInputStream fis = new FileInputStream(new File("C:\\API Testing\\ExcelDataDriven.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		Row row  = sheet.getRow(0);
		int cellCount = row.getLastCellNum();
		Object[][] data = new Object[rowCount][cellCount];
		System.out.println(rowCount+" "+cellCount);
		for(int i=0;i<rowCount;i++)
		{
			row = sheet.getRow(i);
			for(int j=0;j<cellCount;j++)
			{
				Cell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}
		}
		return data;
	}

}
