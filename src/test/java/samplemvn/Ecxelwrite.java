package samplemvn;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Ecxelwrite {

	public static void main(String[] args) throws Throwable {
		File f = new File("C:\\Users\\DELL\\Desktop\\qwerty.xlsx");
		
		FileInputStream f2 = new FileInputStream(f);
		
		
		Workbook w = new XSSFWorkbook(f2);
		Sheet s = w.getSheet("Excel");
		
		for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
			Row r = s.getRow(i);
			
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell cell = r.getCell(j);
				
				int cellType = cell.getCellType();
				
				if (cellType==1) {
					String value = cell.getStringCellValue();
					if (value.equals("hari")) {
						cell.setCellValue("Shyam");
	}
	}
	}
	}
		FileOutputStream f1 = new FileOutputStream(f);
		w.write(f1);
		f1.close();
	}

    }
