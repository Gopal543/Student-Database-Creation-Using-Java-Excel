package Utils;

import java.io.IOException;

public class ExcelUtilsTest {
	
	public static void main(String[] args) throws IOException{
		
		String excelPath ="./data/Testdata.xlsx";
		String sheetName ="Sheet1";
		ExcelUtils excel = new ExcelUtils(excelPath,sheetName);
		
		excel.getRowCount();
		excel.getCellData(1, 1);
		excel.getCellData(1, 2);
		excel.getCellData(1, 3);
		excel.getCellData(1, 4);
		excel.getCellData(1, 5);
		excel.getCellData(1, 6);
		excel.getCellData(1, 7);
		excel.getCellData(1, 8);
		excel.getCellData(1, 9);
		excel.getCellData(1, 10);
		
	}

}
