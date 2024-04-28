package POIPack;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Practice 
{

	public static void main(String[] args) throws IOException 
	{
		String path="C:\\Users\\Adinath Mhetar\\OneDrive\\Desktop\\new.xlsx";
		FileInputStream fis = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		int sheetscount = workbook.getNumberOfSheets();
		for(int i=0;i<sheetscount;i++) 
		{
			if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase("data")) 
			{
				XSSFSheet sheet = workbook.getSheetAt(i);
				Iterator<Row> rows = sheet.rowIterator();
				Row firstrow = rows.next();
				
				Iterator<Cell> cells = firstrow.cellIterator();
				
				
				
				
				
				
				
				
				
			}
			
			
			
			
			
			
			
			
			
			
			
		}
		
		
		
	}

}
