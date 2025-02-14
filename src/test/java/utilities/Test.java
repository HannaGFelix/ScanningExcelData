package utilities;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	
	RetreaveData rData = new RetreaveData();
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public Test(String excelPath, String sheetName) {
		try{
			
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
		
		}catch(Exception exp){
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
		}
	}
	
	public static void main(String[] args) {
		
		String excelPath = "C:\\Users\\hanna\\Desktop\\Github\\ReadData\\src\\test\\resources\\Keystone-online-freedata.xlsx";
		String sheetName = "Sheet1";
		
		Test test = new Test(excelPath,sheetName);
		
		
		int month = 1;
		int year = 3;
		
		RetreaveData rData = new RetreaveData();
		
		DataFormatter formatter = new DataFormatter();
		Object date = formatter.formatCellValue(sheet.getRow(1).getCell(0));
		
		rData.setDate(date);
		rData.setMonth(month);
		rData.setYear(year);
		
		System.out.println(date);
		
		
	}
	
	

}
