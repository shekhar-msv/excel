package test;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.formula.functions.Vlookup;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class VlookUpExample {

	public static void main(String[] args) throws IOException {
		Path pathBook1 = Paths.get("C:\\Users\\Chandra\\Desktop\\TestDataOuput1.xlsx\\");
	    InputStream is = Files.newInputStream(pathBook1);
	    XSSFWorkbook book1 = new XSSFWorkbook(is);Cell cell = book1.getSheetAt(0).createRow(2).createCell(0);
	    Vlookup vs = new Vlookup();
	    vs.evaluate(1, 4, 1, 2, 2, false);
	    cell.setCellFormula("A2+TestDataOuput2.xlsx]Sheet1!A1");
	    FormulaEvaluator mainEvaluator = book1.getCreationHelper().createFormulaEvaluator();
	    XSSFWorkbook book2 = new XSSFWorkbook("C:\\Users\\Chandra\\Desktop\\TestDataOuput2.xlsx");
	    Map<String, FormulaEvaluator> workbooks = new HashMap<String, FormulaEvaluator>();
	    workbooks.put("TestDataOuput1.xlsx", mainEvaluator);
	    workbooks.put("TestDataOuput2.xlsx", book2.getCreationHelper().createFormulaEvaluator());
	    mainEvaluator.setupReferencedWorkbooks(workbooks);
	//  mainEvaluator.evaluateAll();                            // doesn't work.
	//  XSSFFormulaEvaluator.evaluateAllFormulaCells(book1);    // doesn't work.
	    mainEvaluator.evaluateFormulaCell(cell);
	    System.out.println(cell.getNumericCellValue());
	    book2.close();
	    // Close and write workbook 1
	    is.close();
	    OutputStream os = Files.newOutputStream(pathBook1);
	    book1.write(os);
	    os.close();
	    book1.close();
	     
	    
	}
}