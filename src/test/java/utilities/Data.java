package utilities;



public class Data {
	
	public static void main(String[] args) {
		
		String excelPath = "C:\\Users\\hanna\\Desktop\\Github\\ReadData\\src\\test\\resources\\Keystone-online-freedata.xlsx";
		String sheetName = "Sheet1";
		
		ReadData readData = new ReadData(excelPath, sheetName);
		
		ReadData.getData();
		ReadData.getCellData(1,0);
		ReadData.getCellData(1,2);
		
	
	}

}
