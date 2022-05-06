package UtilsLayerDDT;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	
	XSSFWorkbook workbook;
	
	//constructor with 1 String arg
	public ExcelReader(String path)  {
       
		File f = new File(path);
		System.out.println(f.exists());
		
		try {
			FileInputStream fis = new FileInputStream(f);
			 workbook = new XSSFWorkbook(fis);//HSSFWorkbook
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	//non static method with String return type and 3 int args
	public String getDataFromExcel(int sheetIndex,int rowIndex,int columnIndex) {
		return workbook.getSheetAt(sheetIndex).getRow(rowIndex).getCell(columnIndex).getStringCellValue();
	 //XSSFSheet
	}
	
	//non static method with int return type and 1 int arg
	public int countTotleRow(int sheetIndex) {
		return workbook.getSheetAt(sheetIndex).getLastRowNum();
		
	}
	
	//non static method with int return type and 1 int arg
	public int countTotleCell(int sheetIndex) {
		return workbook.getSheetAt(sheetIndex).getRow(0).getLastCellNum();
	}
				
}
