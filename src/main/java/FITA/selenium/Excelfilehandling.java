package FITA.selenium;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Excelfilehandling
{
	
	String filepath = System.getProperty("user.dir")+"//Files//SampleData.xls";

	
	
	public void Readdata() throws IOException
	{
		File F = new File(filepath);
		FileInputStream FS1 = new FileInputStream (F);
		HSSFWorkbook wbk = new HSSFWorkbook(FS1);
		Sheet sheet = wbk.getSheet("SampleData");
		int totalrows = sheet.getPhysicalNumberOfRows();
		String filepath1 = System.getProperty("user.dir")+"//Output//Output.xls";
		File F1= new File(filepath1);
		FileOutputStream FS2 = new FileOutputStream (F1);
		HSSFWorkbook wbk1 = new HSSFWorkbook();
		Sheet sheet1 = wbk1.createSheet("outputData");
		for (int i=0;i<totalrows;i++)
		{
			Row row = sheet.getRow(i);
			int totalColumns = row.getLastCellNum();
			Row row1= sheet1.createRow(i);
			
			for(int j=0;j<totalColumns;j++)
				
			{
				Cell cell= row.getCell(j);
				Cell cell1=row1.createCell(j);
				String a= GetData(cell);
				cell1.setCellValue(a);
			
		    }
			
			wbk1.write(F1);
		//	wbk.close();
			FS1.close();
			FS2.close();
		}
	}
	
		public String  GetData(Cell cellValue)
		{
			switch(cellValue.getCellType())
			{
			case STRING:
				return cellValue.getStringCellValue();
			case NUMERIC:
				DataFormatter DM= new DataFormatter();
				return DM.formatCellValue(cellValue);
			default:
				break;
			}
			return null;
		}
		
		
	
	
	

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		Excelfilehandling E= new Excelfilehandling();
		E.Readdata();
		
		
	}

}
