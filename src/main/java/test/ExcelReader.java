package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class ExcelReader {

	public static void main(String[] args) {
		
		InputStream is = new FileInputStream(new File("C:\\Users\\Chandra\\Desktop\\TestData.xlsx\\"));
		StreamingReader reader = StreamingReader.builder()
		        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
		        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
		        .sheetIndex(0)        // index of sheet to use (defaults to 0)
		        .sheetName("sheet1")  // name of sheet to use (overrides sheetIndex)
		        .read(is);  
		// TODO Auto-generated method stub

	}

}
