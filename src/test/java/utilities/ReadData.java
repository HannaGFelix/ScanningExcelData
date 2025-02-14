package utilities;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {
	
	/*By making the variables static it becomes a global variable and
	can be referenced from any function.*/
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	
	public ReadData(String excelPath, String sheetName){
		
		try{
			
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
		
		}catch(Exception exp){
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
		}
		
	}
	
	public static void getData() {

	int rowCount = sheet.getPhysicalNumberOfRows();
	System.out.println("No of Rows: "+rowCount);
	System.out.println("By Hanna Felix");
	
	
	}
	
	public static void getCellData(int rowNum, int colNum) {
		
		XSSFCell date = sheet.getRow(1).getCell(0);
		Object month = sheet.getRow(1).getCell(1);
		
			
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
		
		
		
		System.out.println(value);
		System.out.println(date);
		System.out.println(month);
		
	}

}
	
	
