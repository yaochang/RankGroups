import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXExcel 
{
	private String filename;
	private XSSFWorkbook workbook;
	private XSSFSheet[] sheets;
	private int sheets_cnt;
	private FileOutputStream fileOut;
	
	public XLSXExcel(String file) {
		try{
			InputStream ExcelFileToRead = new FileInputStream(file);
			filename = file;
			workbook = new XSSFWorkbook(ExcelFileToRead);
			sheets_cnt = workbook.getNumberOfSheets();
			sheets = new XSSFSheet[sheets_cnt];
			for(int i=0;i<sheets_cnt;i++){
				sheets[i] = workbook.getSheetAt(i);
			}
			fileOut = null;
		}catch(Exception e){
			System.out.println(e.toString());
		}
	}
	
	public XLSXExcel() {
		workbook = null;
		sheets = null;
		sheets_cnt = 0;
		filename = null;
		fileOut = null;
	}
	

	public String getContentFromCell(int colid, int rowid, int sheet_id) {
		XSSFRow row = sheets[sheet_id].getRow(rowid);
		XSSFCell cell = row.getCell(colid);
		if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC){
			return cell.getNumericCellValue()+"";
		}else if(cell.getCellType() == XSSFCell.CELL_TYPE_STRING){
			return cell.getStringCellValue();
		}else{
			return null;
		}
	}
	
	public void setContent(int colid, int rowid, int sheet_id, String content) {
		XSSFRow row = null;
		XSSFCell cell = null;
		row = sheets[sheet_id].getRow(rowid);
		if(row == null) row = sheets[sheet_id].createRow(rowid);
		cell = row.getCell(colid);
		if(cell == null) cell = row.createCell(colid);
		cell.setCellValue(content);
	}

	public void createNewExcel(String file) {
		try{
			filename = file;
			workbook = new XSSFWorkbook();
			sheets_cnt = 10; //default
			sheets = new XSSFSheet[sheets_cnt];
			for(int i=0;i<sheets_cnt;i++){
				sheets[i] = workbook.getSheetAt(i);
			}
			fileOut = new FileOutputStream(file);
		}catch(Exception e){
			System.out.println(e.toString());
		}
		
	}
}






