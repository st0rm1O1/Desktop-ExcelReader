package com.github.st0rm1O1.common;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class FileIO {
	
	private FileInputStream fileInput;
	private FileOutputStream fileOutput;
	private XSSFWorkbook workBook;
	private XSSFSheet spreadSheet;
	
	
	
	public void createExcel(String documentPath, String fileName) {
		
		try {
		
			if (fileName != null) {
				
				workBook = new XSSFWorkbook();
				spreadSheet = workBook.createSheet("Kunal");
				spreadSheet.createRow(0).createCell(0).setCellValue("NO DATA INSERTED!");
				
				fileOutput = new FileOutputStream(documentPath + File.separator + fileName + ".xlsx");
				workBook.write(fileOutput);
				
				close();
				
			} // if
		
		} // try
		
		catch (Exception e) {
			e.printStackTrace();
		}
		
	} // createExcel()
	
	
	
	
	
	
	public void insertRecordExcel(int startRow, int indexRow, File file) {
		
		try {
		
			fileInput = new FileInputStream(file);
			workBook = new XSSFWorkbook(fileInput);
			spreadSheet = workBook.getSheetAt(0);
			
		    int endRow = spreadSheet.getLastRowNum();
		    if (endRow < startRow) {
		        spreadSheet.createRow(startRow);
		    }
		    
		    if (startRow >= spreadSheet.getLastRowNum()) {
		    	spreadSheet.createRow(spreadSheet.getLastRowNum());
		    }
		    
		    else {
		    	spreadSheet.shiftRows(startRow, endRow, indexRow, true, true);
			    spreadSheet.createRow(startRow);
		    }
		    
		    
		    fileOutput = new FileOutputStream(file.getAbsolutePath());
		    workBook.write(fileOutput);
		    
		    close();
		    
		} // try
		
		catch (Exception e) {
			e.printStackTrace();
		}
		
	} // insertRecordExcel()
	
	
	
	
	
	
	public void updateRecordExcel(int indexRow, int indexCol, File file, String value) {
		
		try {
		
			fileInput = new FileInputStream(file);
			workBook = new XSSFWorkbook(fileInput);
			spreadSheet = workBook.getSheetAt(0);
		    Cell cell = spreadSheet.getRow(indexRow).getCell(indexCol);
		    
		    if (cell == null)
		    	spreadSheet.getRow(indexRow).createCell(indexCol).setCellValue(value);
		    
		    else cell.setCellValue(value);
		    
		    
		    fileOutput = new FileOutputStream(file.getAbsolutePath());
		    workBook.write(fileOutput);
		    
		    close();
		    
		} // try
		
		catch (Exception e) {
			e.printStackTrace();
		}
		
	} // deleteRecordExcel()
	
	
	
	
	
	
	public void deleteRecordExcel(int startRow, int indexRow, File file) {
		
		try {
		
			fileInput = new FileInputStream(file);
			workBook = new XSSFWorkbook(fileInput);
			spreadSheet = workBook.getSheetAt(0);
			
		    int endRow = spreadSheet.getLastRowNum();
		    if (startRow >= 0 && startRow < endRow)
		    	spreadSheet.shiftRows(startRow+1, endRow, -1);
		    
		    if (startRow >= endRow) {
	            Row removingRow = spreadSheet.getRow(startRow);
	            if (removingRow != null) {
	                spreadSheet.removeRow(removingRow);
	            }
	        }
		    
		    
		    fileOutput = new FileOutputStream(file.getAbsolutePath());
		    workBook.write(fileOutput);
		    
		    close();
		    
		} // try
		
		catch (Exception e) {
			e.printStackTrace();
		}
		
	} // deleteRecordExcel()
	
	
	
	
	
	
	protected void close() {
		
		try {
			
			if (fileInput != null)
				fileInput.close();
			
			if (fileOutput != null)
				fileOutput.close();
			
			if (workBook != null)
				workBook.close();
			
		} catch (Exception e) { e.printStackTrace(); }
		
	} // close()

	
} // class