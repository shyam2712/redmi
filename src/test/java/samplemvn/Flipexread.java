package samplemvn;

import java.io.File;
import java.io.FileInputStream;


import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Flipexread {

	public static void main(String[] args)throws Exception {
		// TODO Auto-generated method stub
		File f = new File("C:\\Users\\DELL\\eclipse-workspace\\samplemvn\\target\\redmii.xlsx");
		 
		FileInputStream f1  = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(f1);
		Sheet s = w.getSheet("text,number");
		
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell cell = r.getCell(j);
			     System.out.println(r);
				System.out.println(cell);}}}}
				
			   
	                        
