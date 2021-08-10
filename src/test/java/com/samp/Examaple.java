package com.samp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellAlignment;

public class Examaple {
	public static void main(String[] args) throws IOException {
		
		File file = new File("C:\\Users\\HP\\eclipse-workspace\\DataDriven\\TestData\\RegisterAutomation.xlsx");
		FileInputStream stream=new FileInputStream(file);
		Workbook workbook=new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet("Sample");
		
	 for (int i = 0; i <sheet.getPhysicalNumberOfRows() ; i++) {
		 
		 Row row = sheet.getRow(i);
		 for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
			 Cell cell = row.getCell(j);
			 
			 int cellType = cell.getCellType();
			 if (cellType==1) {
				 
				 String string = cell.getStringCellValue();
				 System.out.println(string);
			 }
			 else if (DateUtil.isCellDateFormatted(cell)) {
				 Date date = cell.getDateCellValue();
				
			}
				 
				 if (cellType==0) {
					 double numericCellValue = cell.getNumericCellValue(); //type casting
					 long l=(long)numericCellValue;
					 String string = String.valueOf(l);
					 System.out.println(string);
					
				}
				 
				
			}
			
		}
	}
		
	}


