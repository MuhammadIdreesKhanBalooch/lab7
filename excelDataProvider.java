package utilities;

import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class excelDataProvider {
		@Test(dataProvider = "test1data")
		public void test1(String firstname, String lname) {
			System.out.println(firstname + " | " + lname );
		}
	
	
		@DataProvider(name = "test1data")
		public Object[][] getData() {
			String excelPath = "C:\\Users\\4073\\eclipse-workspace\\myMaven\\excel\\Book1.xlsx";
			Object data[][] = testData(excelPath, "Sheet1");
			return data;
		}
	
	public Object[][] testData(String excelPath, String sheetName) {
		excelUtilities excel = new excelUtilities( excelPath,  sheetName);
		int rowCount =  excel.getRowCount();
		int colCount = excel.getColCount();
		
		Object data[][] = new Object[rowCount - 1][colCount  ];
		
		for( int i = 1; i < rowCount; i ++ ) {
			for( int j = 0; j < colCount ; j++ ) {
				String cellData = excel.getCellDataString(i, j);
//				System.out.print(cellData + " | ");
				data[i-1][j] = cellData;
			}
//		System.out.println();
		}
		return data; 
	}
}
