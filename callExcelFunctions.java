package utilities;

public class callExcelFunctions {
	public static void main(String[] args) {
		String 		projPath = System.getProperty("user.dir");
		excelUtilities excel = new excelUtilities( projPath + "\\excel\\Book1.xlsx", "Sheet1" );
		excel.getRowCount();
		excel.getColCount(); 
	//	excel.getCellDataString(1, 0);
	
	}
}
