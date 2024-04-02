/**
 * 
 */
/**
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.Iterator;
import java.util.logging.*;
import java.io.FileOutputStream; 
import java.io.IOException; 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row; 
import org.apache.poi.xssf.usermodel.XSSFSheet; 
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 

public class Hello {
	public static void main(String [] args) {
		String filelocation="c:\\January 2024.xlsx";
		FileInputStream file;
		try {
			file = new FileInputStream(new File(filelocation));
			try {
				XSSFWorkbook work = new XSSFWorkbook(file);
				XSSFSheet sheet = work.getSheet("Claim Form");
	            // Iterate through each rows one by one 
	            Iterator<Row> rowIterator = sheet.iterator(); 
	  
	            // Till there is an element condition holds true 
	            while (rowIterator.hasNext()) { 
	  
	                Row row = rowIterator.next(); 
	  
	                // For each row, iterate through all the 
	                // columns 
	                Iterator<Cell> cellIterator 
	                    = row.cellIterator(); 
	  
	                while (cellIterator.hasNext()) { 
	  
	                    Cell cell = cellIterator.next(); 

	                    // Checking the cell type and format 
	                    // accordingly 
	                    switch (cell.getCellType()) { 
	                    
	                    // Case 1 
	                    case NUMERIC: 
	                        System.out.print( 
	                            cell.getNumericCellValue() 
	                            + ":NUMBER"); 
	                        break; 
	  
	                    // Case 2 
	                    case STRING: 
	                        System.out.print( 
	                            cell.getStringCellValue() 
	                            + ":STR"); 
	                        break; 
	                    } 
	                } 
	  
	                System.out.println(""); 
	            } 				
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			// if File is not found
			e.printStackTrace();
		}
	
		
	}
}