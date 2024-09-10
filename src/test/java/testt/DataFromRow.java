package testt;

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

public class DataFromRow {

	//Identify Testcases coloumn by scanning the entire 1st row
	//once coloumn is identified then scan entire testcase coloumn to identify purchase testcase row
	//after you grab purchase testcase row = pull all the data of the data of that row and feed into test
	
	public ArrayList getData(String testcaseName) throws IOException
	{
		ArrayList<String> a = new ArrayList<String>();
		
		File file = new File(System.getProperty("user.dir")+"\\TestData\\rahulsh.xlsx");
		FileInputStream fis = new FileInputStream(file);
		
		XSSFWorkbook workbook = new XSSFWorkbook(fis); 
		//workbook.getSheet("testdata");
		int sheets = workbook.getNumberOfSheets();
		
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				//Identify Testcases coloum by scanning the entire 1st row
				Iterator<Row> rows = sheet.iterator();//sheet is collection or rows
				Row firstrow = rows.next();
				Iterator<Cell> ce = firstrow.cellIterator();//row is collection of cells
				int k=0;
				int coloumn=0;
				while(ce.hasNext())
				{
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("Testcases"))
					{
						coloumn=k;
					}
					k++;
				}
				//once coloumn is identified then scan entire testcase colom to identify purchase testcase row
				while(rows.hasNext())
				{
					Row r = rows.next();
					
					if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcaseName))
					{
						//after you grab purchase testcase row = pull all the data of the data of that row and feed into test
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
		return a;
	}
}
