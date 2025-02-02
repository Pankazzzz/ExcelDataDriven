package ExcelDriven.ExcelDriven;

import java.io.File;
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

public class DataDriven {

	public static ArrayList<String> getData(String testcaseName) throws Exception
	{
		
		ArrayList<String> data = new ArrayList<String>();
		FileInputStream fis = new FileInputStream(new File("C:\\API Testing\\ExcelDataDriven.xlsx"));
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetCount= workbook.getNumberOfSheets();
		for(int i=0;i<sheetCount;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("Data"))
			{
				XSSFSheet sheet  = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.iterator();
				Row r = rows.next();
				Iterator<Cell> cells  = r.cellIterator();
				int k=0;
				int column=0;
				while(cells.hasNext())
				{
					Cell c = cells.next();
					if(c.getStringCellValue().equalsIgnoreCase("Purchase"))
					{
						column=k;
						break;
					}
					k++;
				}
				System.out.println(column);
				while(rows.hasNext())
				{
					Row row = rows.next();
					if(row.getCell(column).getStringCellValue().equalsIgnoreCase(testcaseName))
					{
						Iterator<Cell> dataCells  = row.cellIterator();	
						while(dataCells.hasNext())
						{
							Cell c = dataCells.next();
							if(c.getCellType()==CellType.STRING)
							//System.out.println(dataCells.next().getStringCellValue());
							data.add(c.getStringCellValue());
							else
							{
								data.add(NumberToTextConverter.toText(c.getNumericCellValue()));
							}
						}
					}
					
					
				}
				
			}
		}
		return data;
	}
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		ArrayList data =getData("Pankaj");
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
	}

}
