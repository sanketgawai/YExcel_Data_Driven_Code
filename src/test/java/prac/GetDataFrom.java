package prac;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetDataFrom {

	public static void main(String[] args) throws IOException {
		
		ArrayList<String> a = new ArrayList<String>();
		
		File file = new File(System.getProperty("user.dir")+"\\TestData\\rahulsh.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fis); 
		int sheets =workbook.getNumberOfSheets();
		
		for(int i=0;i<=sheets;i++) 
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet =workbook.getSheetAt(i);
				Iterator<Row> rows =sheet.rowIterator();
				Row first = rows.next();
				Iterator<Cell> ce = first.cellIterator();
				int k=0;
				int column=0;
				while(ce.hasNext())
				{
					Cell value=ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("Testcases"))
					{
						column=k;
					}
					k++;
				}
				while(rows.hasNext())
				{
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Add Profile"))
					{
						Iterator<Cell> cv = r.cellIterator();
						while(cv.hasNext())
						{
							Cell c = cv.next();
							switch(c.getCellType()) {
							
							case STRING:
								a.add(c.getStringCellValue());
								break;
							case NUMERIC:
								a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
								break;
							case BOOLEAN:
								a.add(Boolean.toString(c.getBooleanCellValue()));
								break;
							case FORMULA:
								a.add(c.getCellFormula());
								break;
								default:
									a.add("Unsupported cell type");
									break;
							}
						}
					}
				}
			}
		}
		
	}
	
}
