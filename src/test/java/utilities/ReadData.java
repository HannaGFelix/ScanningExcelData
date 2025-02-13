package utilities;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
	
	public static void getData() {
		
	try {
	
	String excelPath = "C:\\Users\\hanna\\Desktop\\Github\\ReadData\\src\\test\\resources\\Keystone-online-freedata.xlsx";
	XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
	
	XSSFSheet sheet = workbook.getSheet("Sheet1");
	int rowCount = sheet.getPhysicalNumberOfRows();
	
	System.out.println("No of Rows: "+rowCount);
	System.out.println("By Hanna Felix");
	
	
	}catch(Exception exp){
		System.out.println(exp.getCause());
		System.out.println(exp.getMessage());
	}
	
	
	}
	
	public static void getCellData() {
		
		try {
		String excelPath = "C:\\Users\\hanna\\Desktop\\Github\\ReadData\\src\\test\\resources\\Keystone-online-freedata.xlsx";
		XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
		
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		DataFormatter formatter = new DataFormatter();
		formatter.formatCellValue(sheet.getRow(1).getCell(2));
		
		double value = sheet.getRow(1).getCell(2).getNumericCellValue();
		System.out.println(value);
		
		}catch(Exception exp) {
			
		}

	}
}
	
	
