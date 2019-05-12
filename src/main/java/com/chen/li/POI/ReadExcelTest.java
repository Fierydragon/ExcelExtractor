package com.chen.li.POI;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcelTest {
	public static void main(String[] args) throws IOException {
		testPOI();
//		Class<?> type = Integer.class;
//        Object cast = type.cast(Integer.parseInt("1"));
//        
//        System.out.println(cast);
	}

	private static void testPOI() throws IOException, FileNotFoundException {
		String property = System.getProperty("user.dir");
		System.out.println(property);
		File f = new File("SampleSS.xlsx");
	    Workbook wb = WorkbookFactory.create(f);
	    DataFormatter formatter = new DataFormatter();
	    
	    int i = 1;
	    int numberOfSheets = wb.getNumberOfSheets();
	    for ( Sheet sheet : wb ) {
	        System.out.println("Sheet " + i + " of " + numberOfSheets + ": " + sheet.getSheetName());
	        for ( Row row : sheet ) {
	            System.out.println("\tRow " + row.getRowNum());
	            for ( Cell cell : row ) {
	                System.out.println("\t\t"+ cell.getAddress().formatAsString() + ": " + formatter.formatCellValue(cell));
	            }
	        }
	        Row row = sheet.getRow(0);
	        int firstRowNum = sheet.getFirstRowNum();
	        int lastRowNum = sheet.getLastRowNum();
	        System.out.println(String.format("First row num: %d, last row num: %d.", firstRowNum, lastRowNum));
	    }
	    
//	    new ExcelExtractor("SampleSS.xlsx");

	    // Modify the workbook
	    Sheet sh = wb.createSheet("new sheet");
	    Row row = sh.createRow(7);
	    Cell cell = row.createCell(42);
//	    cell.setActiveCell(true);
	    cell.setCellValue("The answer to life, the universe, and everything");

	    // Save and close the workbook
	    OutputStream fos = new FileOutputStream("SampleSS-updated.xlsx");
	    wb.write(fos);
	    fos.close();
	}
}
