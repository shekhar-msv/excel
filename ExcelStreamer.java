package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import com.monitorjbl.xlsx.StreamingReader;

public class ExcelStreamer {

	public static void main(String[] args) throws FileNotFoundException {
		Instant start =Instant.now();
		String filepath = "files\\input\\Google-Playstore-1.xlsx";
		readExcel(filepath);
		Instant finish =Instant.now();
    	long timeElapsed = Duration.between(start, finish).toMillis();
    	System.out.println("Time consumed = " + timeElapsed);
	}
		public static void readExcel(String filepath) throws FileNotFoundException {
		InputStream is = new FileInputStream(new File(filepath));		
		
    	Workbook workbook = StreamingReader.builder()
    			.rowCacheSize(100)
    			.bufferSize(4096)
    			.open(is);
    	Sheet sheet = workbook.getSheetAt(0);
	    List sheetData = new ArrayList();    	   
	    Iterator rows = sheet.rowIterator();
	    while (rows.hasNext()) {
	    	Row r = (Row) rows.next();    	    			
			Iterator cells = r.cellIterator();
			List data = new ArrayList();
				while (cells.hasNext()) {
		          Cell c = (Cell) cells.next();
		        	data.add(c);
		        }
	        sheetData.add(data);
	    }   	    
    	System.out.println(sheetData.size());
    }
}