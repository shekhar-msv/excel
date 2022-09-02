package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataSanitization {

	public static void main(String[] args) throws IOException  {
		File file = new File("C:\\Users\\Chandra\\Desktop\\TestData.xlsx");   
		FileInputStream fip = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fip); 
        XSSFSheet sheet = workbook.getSheetAt(0);
        System.out.println(sheet.getSheetName());
        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(0).getLastCellNum();
        Iterator iterator = sheet.iterator();
        while(iterator.hasNext()) {
        	XSSFRow row =(XSSFRow) iterator.next();
        	Iterator cellIterator = row.cellIterator();
        	while(cellIterator.hasNext()) {
        		XSSFCell cell = (XSSFCell) cellIterator.next();
        		switch(cell.getCellType())
	          	  {
		          	  case 1: System.out.print(cell.getStringCellValue()); break;
		          	  case 0: System.out.print(cell.getNumericCellValue()); break;
	          	  } 
        		System.out.print(" | ");
        	}
        	System.out.println();
        }     
	}
}