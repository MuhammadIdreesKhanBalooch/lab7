
package utilities;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelUtilities {
	
	static String projPath;
	static XSSFWorkbook workBook;
	static XSSFSheet sheet;
	
//constructor
	public excelUtilities(String excelPath, String Sheetname) {
		
			try {
//		projPath = System.getProperty("user.dir");	
		 workBook = new XSSFWorkbook(excelPath);
		 sheet = workBook.getSheet(Sheetname);
			}
			catch(Exception myExc) {
		
				myExc.printStackTrace();
		
			}
		
	}
	
	public static void main(String[] args) {
//		getRowCount();
		getCellDataString(0,0);
	}
	
	public static int getRowCount() {
		int rowCount = 0;	
		try {
//				 projPath = System.getProperty("user.dir");	
//				 workBook = new XSSFWorkbook(projPath + "\\excel\\Book1.xlsx");
//				 sheet = workBook.getSheet("Sheet1");
				
			rowCount =	sheet.getPhysicalNumberOfRows();
					System.out.println("total no. of rows : " + rowCount);
				}
					catch(Exception myExc) {
					System.out.println(myExc.getMessage()); ;
					System.out.println(myExc.getCause());
					myExc.getStackTrace();
			
				}
			return rowCount;
	}
	
	
	public static int getColCount() {
		int colCount = 0;
		try {
//			 projPath = System.getProperty("user.dir");	
//			 workBook = new XSSFWorkbook(projPath + "\\excel\\Book1.xlsx");
//			 sheet = workBook.getSheet("Sheet1");
			colCount =	sheet.getRow(0).getPhysicalNumberOfCells();
				System.out.println("total no. of columns : " + colCount);
			}
				catch(Exception myExc) {
				System.out.println(myExc.getMessage()); ;
				System.out.println(myExc.getCause());
				myExc.getStackTrace();
		
			}
		return colCount;
}
	
	
	
	public static String getCellDataString(int rowCount, int colCount) {
		String cellData = null;	
		try {
//				 projPath = System.getProperty("user.dir");	
//				 workBook = new XSSFWorkbook(projPath + "\\excel\\Book1.xlsx");
//				 sheet = workBook.getSheet("Sheet1");
				cellData =	sheet.getRow(rowCount).getCell(colCount).getStringCellValue();
				 System.out.println(cellData);  
		} 
				catch(Exception myExc) {
					System.out.println(myExc.getMessage()); ;
					System.out.println(myExc.getCause());
					myExc.getStackTrace();
			
				}
		return cellData;	
	}
}
