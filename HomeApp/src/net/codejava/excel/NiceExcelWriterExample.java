package net.codejava.excel;

import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




/**
 * A nice program that writes data to an Excel file in OOP way. 
 * @author www.codejava.net
 *
 */
public class NiceExcelWriterExample {

	public boolean writeExcel(List<Book> listBook, String excelFilePath) throws IOException {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		
		int rowCount = -1;
		
		for (Book aBook : listBook) {
			Row row = sheet.createRow(++rowCount);
			System.out.println("row in sheet :"+row);
			System.out.println("rowCount :"+rowCount);
			writeBook(aBook, row,rowCount);
			
		}
		
		try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
			workbook.write(outputStream);
			return true;
		}		
	}
	
	
	
	private void writeBook(Book aBook, Row row , int rowCount) {
		
	if(rowCount!=0){
		Cell cell = row.createCell(0);
		System.out.println("cell 0"+cell);
		cell.setCellValue(rowCount);
	}else {Cell cell = row.createCell(0);
	System.out.println("cell 0"+cell);
	cell.setCellValue("No");}
	
	
		Cell cell = row.createCell(1);
		System.out.println("cell 1 :"+cell);
		cell.setCellValue(aBook.getTitle());
		String colIndex=CellReference.convertNumToColString(cell.getColumnIndex());
		String rowIndex=CellReference.convertNumToColString(cell.getRowIndex());
		System.out.println("colIndex is :"+colIndex);
		System.out.println("rowIndex is :"+rowIndex);
		System.out.println("------------------------------------");
		/*cell.getSheet().getRow(0).getCell(currentcellIndex)
	    .getRichStringCellValue().toString()*/
		
		cell = row.createCell(2);
		cell.setCellValue(aBook.getAuthor());
		System.out.println("cell 2 :"+cell);
		colIndex=CellReference.convertNumToColString(cell.getColumnIndex());
		rowIndex=CellReference.convertNumToColString(cell.getRowIndex());
		System.out.println("colIndex is :"+colIndex);
		System.out.println("rowIndex is :"+rowIndex);
		System.out.println("------------------------------------");
		
		cell = row.createCell(3);
		cell.setCellValue(aBook.getPrice());
		System.out.println("cell 3 :"+cell);
		colIndex=CellReference.convertNumToColString(cell.getColumnIndex());
		rowIndex=CellReference.convertNumToColString(cell.getRowIndex());
		System.out.println("colIndex is :"+colIndex);
		System.out.println("rowIndex is :"+rowIndex);
		System.out.println("------------------------------------");
	}
	
	
	
	
	//Gettings Books
	private List<Book> getListBook() {
		Book book0 = new Book("TITLE","AUTHOR","PRICE");
		Book book1 = new Book("Head First Java", "Kathy Serria","79");
		Book book2 = new Book("Effective Java", "Joshua Bloch", "36");
		Book book3 = new Book("Clean Code", "Robert Martin", "42");
		Book book4 = new Book("Thinking in Java", "Bruce Eckel"," 35");
		
		List<Book> listBook = Arrays.asList(book0,book1, book2, book3, book4);
		
		return listBook;
	}
	
	
	
	
	
	
	
	
	
	
	
	
//Main Method 	
	public static void main(String[] args) throws IOException {
		NiceExcelWriterExample excelWriter = new NiceExcelWriterExample();
		List<Book> listBook = excelWriter.getListBook();
		
		Date date = new Date();
		Format formatter = new SimpleDateFormat("YYYY-MM-dd_hh-mm-ss");
	    
		
		String fileName="NiceJavaBooks_"+ formatter.format(date) + ".xls";
		String excelFilePath = "C:\\Local_Drive\\Work\\Rough\\Test\\"+fileName;
		//font start
		
		
		//font last
		if(excelWriter.writeExcel(listBook, excelFilePath)){
			System.out.println("Excel Created Successfully");
	}
		else{System.out.println("Not Created Excel");}
	}

}


