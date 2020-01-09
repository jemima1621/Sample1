package org.cts;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileWrite {
	public static void main(String[] args) throws IOException {
		File loc=new File("C:\\Users\\jemima j\\Desktop\\java programs\\Excel\\Write.xlsx");
		Workbook w=new XSSFWorkbook();
		Sheet s=w.createSheet("greens");
		Row r = s.createRow(5);
		Cell c = r.createCell(6);
		c.setCellValue("Jemima");
		FileOutputStream out=new FileOutputStream(loc);
		w.write(out);
		System.out.println("Write sucessfully");
		
	}

}
