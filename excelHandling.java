package automationFramework;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class excelHandling {

	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private HSSFWorkbook workbook = null;
	private HSSFSheet sheet = null;
	private HSSFRow row = null;
	private HSSFCell cell = null;
	String path = null;
	
	// Create Constructor
	public excelHandling() throws IOException{
		path = System.getProperty("user.dir") + "\\testdata\\testdata.xlsx";
		fis = new FileInputStream(path);
		workbook = new HSSFWorkbook(fis);
		sheet = workbook.getSheetAt(0);
	}
	
	// Provide total number of rows in sheet - testcase
	public int getSheetRows(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		
		return(sheet.getLastRowNum() + 1);
	}
	
	// Provide total number of columns in sheet - testcase
	public int getSheetColumns(String sheetName){
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		
		row = sheet.getRow(0);
		return(row.getLastCellNum());
	}
	
	// Provide cell value - testdata
	public String getCellData(String sheetName, int colNum, int rowNum){
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		
		row = sheet.getRow(rowNum);
		cell = row.createCell(colNum);
		return(cell.getStringCellValue());
	}
	
	// Provide cell value - testdata
	public String getCellData(String sheetName, String colName, int rowNum){
		int colNum = -1;
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		
		for(int i=0; i<getSheetColumns(sheetName); i++){
			row = sheet.getRow(0);
			cell = row.getCell(i);
			
			if(cell.getStringCellValue().equals(colName)){
				colNum = cell.getColumnIndex();
				break;
			}
		}
		row = sheet.getRow(rowNum);
		cell = row.getCell(colNum);
		return(cell.getStringCellValue());
	}
	
	// To set a cell data - testcase
	public void setCellData(String sheetName, int colNum, int rowNum, String str){
		int index = workbook.getSheetIndex(sheetName);
		sheet = workbook.getSheetAt(index);
		
		row = sheet.getRow(rowNum);
		cell = row.createCell(colNum);
		cell.setCellValue(str);
		
		try {
			fileOut = new FileOutputStream(path);
			try {
				workbook.write(fileOut);
				fileOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
	public static void main(String[] args) throws IOException{
		excelHandling reader = new excelHandling();
//		System.out.println(reader.getSheetRows("LoginTest"));
//		System.out.println(reader.getSheetRows("SignUpTest"));
//		System.out.println(reader.getSheetColumns("LoginTest"));
//		System.out.println(reader.getSheetColumns("SignUpTest"));
//		System.out.println(reader.getCellData("LoginTest", 1, 1));
//		System.out.println(reader.getCellData("SignUpTest", 1, 1));
//		System.out.println(reader.getCellData("LoginTest", "password", 1));
//		System.out.println(reader.getCellData("SignUpTest", "lastname", 2));
		reader.setCellData("LoginTest", 1, 1, "Omkar");
		reader.setCellData("SignUpTest", 2, 1, "Tejas");


	}
	

	
}
