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

public class XLSExcel 
{
	private String filename;
	private HSSFWorkbook workbook;
	private HSSFSheet[] sheets;
	private int sheets_cnt;
	private FileOutputStream fileOut;
	
	public XLSExcel(String file) {
		try{
			InputStream ExcelFileToRead = new FileInputStream(file);
			filename = file;
			workbook = new HSSFWorkbook(ExcelFileToRead);
			sheets_cnt = workbook.getNumberOfSheets();
			sheets = new HSSFSheet[sheets_cnt];
			for(int i=0;i<sheets_cnt;i++){
				sheets[i] = workbook.getSheetAt(i);
			}
			fileOut = null;
		}catch(Exception e){
			System.out.println(e.toString());
		}
	}
	
	public XLSExcel() {
		workbook = null;
		sheets = null;
		sheets_cnt = 0;
		filename = null;
		fileOut = null;
	}
	

	public String getContentFromCell(int colid, int rowid, int sheet_id) {
		HSSFRow row = sheets[sheet_id].getRow(rowid);
		HSSFCell cell = row.getCell(colid);
		if(cell == null) System.out.println("null error");
	//	System.out.println(""+cell.toString());
		if(cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC){
			return cell.getNumericCellValue()+"";
		}else if(cell.getCellType() == HSSFCell.CELL_TYPE_STRING){
			return cell.getStringCellValue();
		}else if(cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
			return cell.getCellFormula();
		}else if(cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
			return null;
		}else if(cell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN){
			return cell.getBooleanCellValue()+"";
		}else if(cell.getCellType() == HSSFCell.CELL_TYPE_ERROR){
			return cell.getErrorCellValue() + "";
		}else{
			return cell.getDateCellValue().toString();
		}
	}
	
	public void setContent(int colid, int rowid, int sheet_id, String content) {
		HSSFRow row = null;
		HSSFCell cell = null;
		row = sheets[sheet_id].getRow(rowid);
		if(row == null) row = sheets[sheet_id].createRow(rowid);
		cell = row.getCell(colid);
		if(cell == null) cell = row.createCell(colid);
		cell.setCellValue(content);
	}

	public void createNewExcel(String file) {
		try{
			filename = file;
			workbook = new HSSFWorkbook();
			sheets_cnt = 10; //default
			sheets = new HSSFSheet[sheets_cnt];
			for(int i=0;i<sheets_cnt;i++){
				sheets[i] = workbook.createSheet();
			}
			fileOut = new FileOutputStream(file);
		}catch(Exception e){
			System.out.println(e.toString());
		}
		
	}
	
	public void saveFile() {
		try{
			workbook.write(fileOut);
			fileOut.flush();
			fileOut.close();
		}catch(Exception e){
			
		}
	}
	
	public void outputall() {
		Iterator rowit = sheets[0].rowIterator();
		while( rowit.hasNext()){
			HSSFRow row = (HSSFRow) rowit.next();
			Iterator cellit = row.cellIterator();
			while(cellit.hasNext()){
				HSSFCell cell = (HSSFCell) cellit.next();
				System.out.println(cell.toString());
				
			}
		}
	}
}






