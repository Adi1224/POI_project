package POIPack;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIDemo 
{
	public static void main(String[] args) throws IOException 
	{
		XSSFWorkbook workbook= new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("EmployeeData");
		
		
//		By using ArrayList object and For each loop
		
		ArrayList<Object[]> empdata = new ArrayList<Object[]>();
		empdata.add(new Object[] {"empid","name","job"});
		empdata.add(new Object[] {"101","poonam","CA"});
		empdata.add(new Object[] {"102","ankesh","guitarist"});
		empdata.add(new Object[] {"103","amruta","police"});
		
		int rowcont=0;
		
		for(Object[]emp:empdata) 
		{
			XSSFRow row = sheet.createRow(rowcont++);
			int colcont=0;
			
			for(Object value:emp) 
			{
				XSSFCell cell = row.createCell(colcont++);
				
			if(value instanceof String)
				cell.setCellValue((String)value);
			if(value instanceof Integer)
				cell.setCellValue((Integer)value);
			if(value instanceof Boolean)
				cell.setCellValue((Boolean)value);
				
				
			}
		}
		

		FileOutputStream fos = new FileOutputStream(".\\Files\\EmployeeData1.xlsx");
		workbook.write(fos);
		System.out.println("file has been created");
		
		
		
//		By using two dimensional array and for loop
		
		
		
		Object[][] empdata2= 
			{
				{"empid","name","job"},
				{101,"Ankesh","Guitarist"},
				{102,"Amruta","Police"},
				{103,"Poonam","CA"}	
			};	
										
		int rownum = empdata2.length;
		int colnum = empdata2[0].length;
		System.out.println(rownum+"  "+colnum);
		
		for(int r=0;r<rownum;r++) 
		{
			XSSFRow row = sheet.createRow(r);
			
			for(int c=0;c<colnum;c++) 
			{
				XSSFCell cell = row.createCell(c);
				Object value=empdata2[r][c];
				
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			
			}
			
		}
		
		FileOutputStream fos1 = new FileOutputStream(".\\Files\\EmployeeData2.xlsx");
		workbook.write(fos1);
		System.out.println("file has been created");
		
	}
}