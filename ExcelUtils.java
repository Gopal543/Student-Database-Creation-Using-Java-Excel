package Utils;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;

public class ExcelUtils { 
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
 
	public ExcelUtils(String excelPath,String sheetName) {
			try{
				
				
				
				workbook =new XSSFWorkbook(excelPath);
			    sheet = workbook.getSheet(sheetName);
			

			}catch(Exception exp){
				System.out.println(exp.getCause());
				System.out.println(exp.getMessage());
				exp.printStackTrace();
			}
		}

	
	
	public static void getCellData(int rowNum,int colNum) throws IOException {
		
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));		
		System.out.println(value);
	}		
	
	public static void getRowCount() {	
		
			
			int rowCount=sheet.getPhysicalNumberOfRows();
			System.out.println("No of Rows:"+rowCount);

		
		}

	}
