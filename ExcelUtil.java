import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.formula.eval.BlankEval;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {

	public HashMap<Integer,ArrayList<String>> getData(String excelPath, String sheetName) throws IOException
	
	{
		FileInputStream fis = new FileInputStream(excelPath);
		XSSFWorkbook workbook = new	XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet(sheetName);
		

		int rows = sheet.getLastRowNum();
		
		HashMap<Integer,ArrayList<String>> data = new HashMap<Integer,ArrayList<String>>();
		
		//Reading data from excel to HashMap
		ArrayList<String> values = null;
		for(int r=0; r<=rows; r++)
		{
			XSSFRow row = sheet.getRow(r);
			values = new ArrayList<String>();
			// loop in the row to get all cells values
			int cellNum = row.getLastCellNum();
			
			for(int i=0; i<cellNum; i++)
			{
				XSSFCell cell = row.getCell(i);
				
				if (cell == null) 
				{
					values.add("");
				} 
				else if(cell.getCellType() == CellType.BLANK)
				{
					values.add("");
				}
				else if(cell.getCellType() == CellType.STRING)
				{ //String cell value
					String value = cell.getStringCellValue();
					values.add(value);
				}
				
				//Numeric cell Value
				else if(cell.getCellType() ==CellType.NUMERIC)
				{
					values.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
				}
				
				//Boolean cell value
				else if(cell.getCellType() ==CellType.BOOLEAN)
				{
					values.add(Boolean.toString(cell.getBooleanCellValue()));
				}
				
				
				
			}
			data.put(r,values);
			
		}
		
		//Read data from HashMap
		
//		for(Map.Entry entry:data.entrySet())
//		{
//			System.out.println(entry.getKey()+" 	" + entry.getValue());
//		}
		workbook.close();
		fis.close();
		return data;
		
	}
	
	
	public String getValue(HashMap<Integer,ArrayList<String>> data, Integer rowNum, String colName) {
		String value = null;
		
		
		int index = data.get(0).indexOf(colName);
		value = data.get(rowNum).get(index);
		
		return value;
		
	}
	
	public void writeValue(HashMap<Integer,ArrayList<String>> data, String cellValue, int rowNum, String sheetName, String colName, String excelPath) throws Exception {
		FileOutputStream fos = null;
		XSSFWorkbook workbook =null;
		
		try {
			FileInputStream fis = new FileInputStream(excelPath);
			
			workbook= new XSSFWorkbook(fis);
			
			XSSFSheet sheet = workbook.getSheet(sheetName);
			
			int index = data.get(0).indexOf(colName);
			
			sheet.getRow(rowNum).createCell(index).setCellValue(cellValue);
			
			fos= new FileOutputStream(excelPath);
			
			workbook.write(fos);
		} catch (Exception e) {
			//throw e;
		} finally {
			if (workbook!= null) workbook.close();
			
			if (fos!=null) fos.close();
		}

		
	}
	
	
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

	
		
	}

}

