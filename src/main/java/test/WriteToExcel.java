package test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Emp Info");
		Object empdata[][] = {  {"EmpID","Name","Job"},
								{101,"Chandra1","Engineer1"},
								{102,"Chandra2","Engineer2"},
								{103,"Chandra3","Engineer3"}
		                     };
		int rowCount =0;
		for(Object emp[]:empdata)
		{
			XSSFRow row = sheet.createRow(rowCount++);
			int colCount =0;
			for(Object value:emp) 
			 {
				XSSFCell cell = row.createCell(colCount++);
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
			 }
		}
		String filepath ="C:\\Users\\Chandra\\Desktop\\TestDataOuput.xlsx";
		FileOutputStream outstream = new FileOutputStream(filepath);
		workbook.write(outstream);
		outstream.close();
		System.out.println("File written succesfully");
	}
}