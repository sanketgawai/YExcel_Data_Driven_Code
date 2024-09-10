package testt;

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

public class DataFromColumn {

	// Method to get data from a specific column
    public ArrayList<String> getDataFromColumn(String columnName) throws IOException {
        ArrayList<String> columnData = new ArrayList<>();

        File file = new File(System.getProperty("user.dir") + "\\TestData\\rahulsh.xlsx");
        FileInputStream fis = new FileInputStream(file);
        
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        int sheets = workbook.getNumberOfSheets();

        for (int i = 0; i < sheets; i++) {
            if (workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
                XSSFSheet sheet = workbook.getSheetAt(i);
                
                // Identify column by scanning the entire first row
                Iterator<Row> rows = sheet.iterator();
                Row firstRow = rows.next();
                Iterator<Cell> cellIterator = firstRow.cellIterator();
                
                int targetColumnIndex = -1;
                int currentIndex = 0;
                
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                        targetColumnIndex = currentIndex;
                        break;
                    }
                    currentIndex++;
                }
                
                if (targetColumnIndex == -1) {
                    throw new IllegalArgumentException("Column name not found: " + columnName);
                }
                
                // Collect data from the identified column
                while (rows.hasNext()) {
                    Row row = rows.next();
                    Cell cell = row.getCell(targetColumnIndex);
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                columnData.add(cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                columnData.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
                                break;
                            case BOOLEAN:
                                columnData.add(Boolean.toString(cell.getBooleanCellValue()));
                                break;
                            case FORMULA:
                                columnData.add(cell.getCellFormula());
                                break;
                            default:
                                columnData.add("Unsupported cell type");
                                break;
                        }
                    } else {
                        columnData.add(""); // Handle cases where cell is null
                    }
                }
            }
        }
        
        workbook.close();
        fis.close();
        return columnData;
    }
}

