package com.leon.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelMain {
	static String path_win = "d:/leo_workbook.xls";
	static String path_linux = "/home/leo/leo_workbook.xls";
	public static void main(String[] args){
		try {
			List<String> me = ReadExcelFile.readit(path_win);
			System.out.println(me.size());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
//		try {
//			createWorkbook(path_win);
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
 finally {
}
	}
	public static HSSFWorkbook createWorkbook(String path) throws IOException{
		Workbook wb = new HSSFWorkbook();
		
	    FileOutputStream fileOut = new FileOutputStream(path);
	    wb.write(fileOut);
	    fileOut.close();
		return null;
	}
	public static HSSFSheet createSheet(HSSFWorkbook wb,String sheetName) {
		HSSFSheet sheet = wb.createSheet(sheetName);
		return sheet;
	}
	
}
