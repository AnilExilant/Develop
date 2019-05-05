package utility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperation {

	 FileInputStream fis;
	 FileOutputStream fos;
	 XSSFWorkbook wb;
	 XSSFSheet sheet;
	 Row r;
	 Cell c;
	public int getRowCount(String xlPath,String sheetName) throws IOException {
		fis=new FileInputStream(xlPath);
		wb=new XSSFWorkbook(fis);
		sheet=wb.getSheet(sheetName);
		int rowCount=wb.getSheet(sheetName).getLastRowNum();
		return rowCount;
	}
	public int getCellCount(String xlPath,String sheetName,int rowNo ) throws IOException {
		fis=new FileInputStream(xlPath);
		wb=new XSSFWorkbook(fis);
		sheet=wb.getSheet(sheetName);
		r=sheet.getRow(rowNo);
		int cellCount=r.getLastCellNum();
		return cellCount;
	}
	
	public String getCellData(String xlPath,String sheetName,int rowNo,int colNo) throws IOException {
		fis=new FileInputStream(xlPath);
		wb=new XSSFWorkbook(fis);
		sheet=wb.getSheet(sheetName);
		r=sheet.getRow(rowNo);
		c=r.getCell(colNo);
		String data="";
		return data;
	}
	public void setCellData(String xlPath,String sheetName,int rowNo,int colNo,String data) throws IOException {
		fis=new FileInputStream(xlPath);
		wb=new XSSFWorkbook(fis);
		sheet=wb.getSheet(sheetName);
		r=sheet.getRow(rowNo);
		c=r.createCell(colNo);
		c.setCellValue(data);
		fos=new FileOutputStream("xlPath");
		wb.write(fos);
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
