package Utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UseParameter {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File excel = new File("C:\\Users\\Nitin\\Downloads\\TE GFM Attendance Format for AY 2019-20 Sem 1.xlsx");
		FileInputStream input = new FileInputStream(excel);
		
		XSSFWorkbook take = new XSSFWorkbook(input);
		System.out.println();
		XSSFSheet getSheet = take.getSheet("NBN");
		System.out.println(getSheet.getPhysicalNumberOfRows());
		System.out.println(getSheet.getActiveCell());
	}

}
