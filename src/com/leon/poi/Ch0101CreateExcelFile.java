package com.leon.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.Region;


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
	private static SimpleDateFormat SF = new SimpleDateFormat("yyyy-MM-DD");
	
	public static void main(String args[]) throws IOException, ParseException{
		createWorkbook(PATH_WIN,FILENAME,FORMAT03);
	}
	/**
	 * Create an excel file in format of excel97-2003 with extension of 'xls'
	 * @param path
	 * @param name
	 * @param format
	 * @return
	 * @throws IOException
	 * @throws ParseException
	 */
	public static HSSFWorkbook createWorkbook(String path,String name,String format) throws IOException, ParseException{
		//create the workbook
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
		sheet_student.addMergedRegion(new Region(0, (short)0, 0, (short)5));
		
		//define the style of a cell
		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
	    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
	    
		Row tempRow = sheet_student.createRow(1);
		
		tempRow.createCell(0).setCellValue("ID");
		tempRow.createCell(1).setCellValue("NAME");
		tempRow.createCell(2).setCellValue("STUDENT_NUMBER");
		tempRow.createCell(3).setCellValue("GRADE");
		tempRow.createCell(4).setCellValue("GENDER");
		tempRow.createCell(5).setCellValue("BIRTHDAY");
		
		tempRow = sheet_student.createRow(2);
		tempRow.createCell(0).setCellValue(1);
		tempRow.createCell(1).setCellValue("Allen");
		tempRow.createCell(2).setCellValue(1201);
		tempRow.createCell(3).setCellValue(3);
		tempRow.createCell(4).setCellValue("M");
		tempRow.createCell(5).setCellValue(SF.parse("1992-08-28"));
		tempRow.getCell(5).setCellStyle(cellStyle);
		
		tempRow = sheet_student.createRow(3);
		tempRow.createCell(0).setCellValue(2);
		tempRow.createCell(1).setCellValue("Bob");
		tempRow.createCell(2).setCellValue(1106);
		tempRow.createCell(3).setCellValue(4);
		tempRow.createCell(4).setCellValue("M");
		tempRow.createCell(5).setCellValue(SF.parse("1991-12-02"));
		tempRow.getCell(5).setCellStyle(cellStyle);
		
		tempRow = sheet_student.createRow(4);
		tempRow.createCell(0).setCellValue(3);
		tempRow.createCell(1).setCellValue("Julianne");
		tempRow.createCell(2).setCellValue(1318);
		tempRow.createCell(3).setCellValue(2);
		tempRow.createCell(4).setCellValue("F");
		tempRow.createCell(5).setCellValue(SF.parse("1993-01-28"));
		tempRow.getCell(5).setCellStyle(cellStyle);
		
		tempRow = sheet_student.createRow(5);
		tempRow.createCell(0).setCellValue(4);
		tempRow.createCell(1).setCellValue("Thomas");
		tempRow.createCell(2).setCellValue(1202);
		tempRow.createCell(3).setCellValue(3);
		tempRow.createCell(4).setCellValue("M");
		tempRow.createCell(5).setCellValue(SF.parse("1993-05-16"));
		tempRow.getCell(5).setCellStyle(cellStyle);
		
		//save the content to an excel file
		FileOutputStream fileOut = new FileOutputStream(path+name+"."+format);
		wb.write(fileOut);
	    fileOut.close();
		
		return null;
	}
	
	public static Row fillRow(HSSFSheet sheet,int rowNum){
		return null;
	}
}
