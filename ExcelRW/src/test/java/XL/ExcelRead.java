package XL;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelRead {
	


	public static void main(String[] args) throws IOException {
		FileInputStream file = null;
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		try
        {
            file = new FileInputStream(new File("src/test/resources/ObsqueraStudents.xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                 
                while (cellIterator.hasNext()) 
                {
                    Cell cell = cellIterator.next();
                   
                    switch(cell.getCellType())
                    {
                    case STRING:
                    	System.out.print(cell.getStringCellValue());
                    	break;
                    case NUMERIC:
                    	System.out.print(cell.getNumericCellValue());
                    	break;
                    	default:
                    		break;
                    }
                    
                    
                    
                    
                }
                System.out.println("");
            }
            
          
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
		finally
		{
			file.close();
		}



	}

}
