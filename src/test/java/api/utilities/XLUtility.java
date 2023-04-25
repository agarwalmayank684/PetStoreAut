package api.utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtility {

	public FileInputStream fi;
	public FileOutputStream fo;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet1;
	public XSSFRow row1;
	public XSSFCell cell1;
	public CellStyle style;
	String path;
	
	public XLUtility(String path) {
		this.path = path;
	}
	
	public int getRowCount(String sheetName) throws IOException {
		
		fi = new FileInputStream(path);
	    workbook = new XSSFWorkbook(fi);
	    sheet1 = workbook.getSheet(sheetName);
	    //XSSFRow row = sheet.getRow(0);
	    int rowCount = sheet1.getLastRowNum();
	    System.out.println("Total Number of Rows in the excel is : "+rowCount);
	    workbook.close();
	    fi.close();
	    return rowCount;
	}
	
	public int getCellCount(String sheetName, int rownum) throws IOException {
		
		fi = new FileInputStream(path);
	    workbook = new XSSFWorkbook(fi);
	    sheet1 = workbook.getSheet(sheetName);
	    row1 = sheet1.getRow(rownum);
	    int cellcount = row1.getLastCellNum();
	    System.out.println("Total Number of Rows in the excel is : "+cellcount);
	    workbook.close();
	    fi.close();
	    return cellcount;
	}
	
	public String getCellData(String sheetName, int rownum, int colnum) throws IOException {
		fi = new FileInputStream(path);
	    workbook = new XSSFWorkbook(fi);
	    sheet1 = workbook.getSheet(sheetName);
	    row1 = sheet1.getRow(rownum);
	    cell1 = row1.getCell(colnum);
	    
	    DataFormatter formatter = new DataFormatter();
	    String data;
	    try{
	    	data = formatter.formatCellValue(cell1);
	    }
	    catch(Exception e) {
	    	data = "";
	    }
	    workbook.close();
	    fi.close();
	    return data;
	}
	
	
}
