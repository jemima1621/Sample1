package org.cts;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileRead {
	public static void main(String[] args) throws IOException {
		File loc=new File("E:\\Eclipse\\cts\\testDatas\\Sample.xlsx");
		FileInputStream stream=new FileInputStream(loc);
		Workbook w=new XSSFWorkbook(stream);
		Sheet s = w.getSheet("Sheet1");
//		Row r = s.getRow(0);
//		Cell c = r.getCell(0);
//		System.out.println(c);
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				System.out.println(c);
				
			}
			
		}
		
		
	}

}
