package utilities;

public class excelUtilities {
	public static void getRowCount() {
		try {
//		String projPath = System.setProperty("user.dir");	
		XSSFWorkbook workBook = new XSSFWorkbook("\\add location of sheet");
		XSSFSheet sheet = workbook.getsheet("Sheet1");
		int rowCount =	sheet.getPhysicalNumberOfRows();
		System.out.println("total no. of rows : " + rowCount);
		}catch(Exception myExc) {
			System.out.println(myExc.getMessage()); ;
			System.out.println(myExc.getCause());
			myExc.getStackTrace();
		
		}
		
	}
}
