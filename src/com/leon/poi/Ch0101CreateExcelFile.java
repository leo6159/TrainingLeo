package com.leon.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.Region;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * 2012-8-28
 * @author leo
 * 
 * <br /> Generate the student table in document/1.1.Get_Start
 * <br />  or https://github.com/leo6159/TrainingLeo/blob/master/document/POI/1.About_Parse_Excel
 * <br /> 
 * <br /> ref-url : http://poi.apache.org/spreadsheet/quick-guide.html
 */
public class Ch0101CreateExcelFile {
	private static String PATH_WIN = "d:/";
	private static String PATH_LINUX = "/home/leo/";
	private static String FILENAME = "leo_workbook";
	private static String FORMAT03 = "xls";
	private static String FORMAT07 = "xlsx";
	private static SimpleDateFormat SF = new SimpleDateFormat("yyyy-MM-dd");
	
	
	public static void main(String args[]) throws IOException, ParseException{
		String formatString = FORMAT03;
		if(formatString.equals(FORMAT03)){
			createWorkbook03(PATH_WIN,FILENAME,FORMAT03);
		}else{
			createWorkbook07(PATH_WIN,FILENAME,FORMAT07);
		}
	}
	
	public static XSSFWorkbook createWorkbook03(String path,String name,String format) throws IOException, ParseException{
		HSSFWorkbook wb = new HSSFWorkbook();
		
		//create the sheets, the sheet name must not exceed 31 characters
		String sheet_student_name = "Student";
		//use the util to replaces invalid characters with a space
		sheet_student_name = org.apache.poi.ss.util.WorkbookUtil.createSafeSheetName(sheet_student_name);
		HSSFSheet sheet_student = wb.createSheet(sheet_student_name);
		//create the title line
		Row rows_student = sheet_student.createRow(0);
		Cell cell_title = rows_student.createCell(0);
		cell_title.setCellValue("Table Of Students");
		//merge cells
		sheet_student.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
		
		//define the style of a cell
		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
	    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
	    
	    ArrayList<Map> data = new ArrayList<Map>();
		try {
			data = getData();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    
		
	    Row tempRow = sheet_student.createRow(1);
		
		tempRow.createCell(0).setCellValue("ID");
		tempRow.createCell(1).setCellValue("NAME");
		tempRow.createCell(2).setCellValue("STUDENT_NUMBER");
		tempRow.createCell(3).setCellValue("GRADE");
		tempRow.createCell(4).setCellValue("GENDER");
		tempRow.createCell(5).setCellValue("BIRTHDAY");
		
	    for(int i=0;i<data.size();i++){
	    	tempRow = sheet_student.createRow(i+2);
			tempRow.createCell(0).setCellValue((Integer) data.get(i).get("ID"));
			tempRow.createCell(1).setCellValue((String) data.get(i).get("NAME"));
			tempRow.createCell(2).setCellValue((String) data.get(i).get("STUDENT_NUMBER"));
			tempRow.createCell(3).setCellValue((Integer) data.get(i).get("GRADE"));
			tempRow.createCell(4).setCellValue((String) data.get(i).get("GENDER"));
			tempRow.createCell(5).setCellValue((Date)(data.get(i).get("BIRTHDAY")));
			tempRow.getCell(5).setCellStyle(cellStyle);
	    }
		
		//save the content to an excel file
		FileOutputStream fileOut = new FileOutputStream(path+name+"."+format);
		wb.write(fileOut);
	    fileOut.close();
		
		return null;
	}
	
	public static XSSFWorkbook createWorkbook07(String path,String name,String format) throws IOException, ParseException{
		XSSFWorkbook wb = new XSSFWorkbook();
		
		//create the sheets, the sheet name must not exceed 31 characters
		String sheet_student_name = "Student";
		//use the util to replaces invalid characters with a space
		sheet_student_name = org.apache.poi.ss.util.WorkbookUtil.createSafeSheetName(sheet_student_name);
		XSSFSheet sheet_student = wb.createSheet(sheet_student_name);
		//create the title line
		Row rows_student = sheet_student.createRow(0);
		Cell cell_title = rows_student.createCell(0);
		cell_title.setCellValue("Table Of Students");
		//merge cells
		sheet_student.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
		
		//define the style of a cell
		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
	    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
	    
	    ArrayList<Map> data = new ArrayList<Map>();
		try {
			data = getData();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	    
		
	    Row tempRow = sheet_student.createRow(1);
		
		tempRow.createCell(0).setCellValue("ID");
		tempRow.createCell(1).setCellValue("NAME");
		tempRow.createCell(2).setCellValue("STUDENT_NUMBER");
		tempRow.createCell(3).setCellValue("GRADE");
		tempRow.createCell(4).setCellValue("GENDER");
		tempRow.createCell(5).setCellValue("BIRTHDAY");
		
	    for(int i=0;i<data.size();i++){
	    	tempRow = sheet_student.createRow(i+2);
			tempRow.createCell(0).setCellValue((Integer) data.get(i).get("ID"));
			tempRow.createCell(1).setCellValue((String) data.get(i).get("NAME"));
			tempRow.createCell(2).setCellValue((String) data.get(i).get("STUDENT_NUMBER"));
			tempRow.createCell(3).setCellValue((Integer) data.get(i).get("GRADE"));
			tempRow.createCell(4).setCellValue((String) data.get(i).get("GENDER"));
			tempRow.createCell(5).setCellValue((Date)(data.get(i).get("BIRTHDAY")));
			tempRow.getCell(5).setCellStyle(cellStyle);
	    }
		
		//save the content to an excel file
		FileOutputStream fileOut = new FileOutputStream(path+name+"."+format);
		wb.write(fileOut);
	    fileOut.close();
		
		return null;
	}
	
	public static ArrayList<Map> getData() throws Exception{
		ArrayList<Map> data = new ArrayList<Map>();
	    Map map = new HashMap();
	    map.put("ID",1);
	    map.put("NAME","Allen");
	    map.put("STUDENT_NUMBER","1201");
	    map.put("GRADE",3);
	    map.put("GENDER","M");
	    map.put("BIRTHDAY",SF.parse("1992-08-28"));
	    data.add(map);
	    
	    map = new HashMap();
	    map.put("ID",2);
	    map.put("NAME","Bob");
	    map.put("STUDENT_NUMBER","1106");
	    map.put("GRADE",4);
	    map.put("GENDER","M");
	    map.put("BIRTHDAY",SF.parse("1991-12-02"));
	    data.add(map);
	    
	    map = new HashMap();
	    map.put("ID",3);
	    map.put("NAME","Julianne");
	    map.put("STUDENT_NUMBER","1318");
	    map.put("GRADE",2);
	    map.put("GENDER","F");
	    map.put("BIRTHDAY",SF.parse("1993-01-28"));
	    data.add(map);
	    
	    map = new HashMap();
	    map.put("ID",4);
	    map.put("NAME","Thomas");
	    map.put("STUDENT_NUMBER","1202");
	    map.put("GRADE",3);
	    map.put("GENDER","M");
	    map.put("BIRTHDAY",SF.parse("1993-05-1"));
	    data.add(map);
	    
	    return data;
	}
	
}
