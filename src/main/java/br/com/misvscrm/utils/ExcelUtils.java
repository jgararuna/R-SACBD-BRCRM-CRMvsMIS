package br.com.misvscrm.utils;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.aspose.cells.FindOptions;
import com.aspose.cells.LookAtType;
import com.spire.xls.CellRange;
import com.spire.xls.Worksheet;
import com.spire.xls.core.converter.spreadsheet.RangeCollection;

import br.com.misvscrm.XlsxVsXls;
import br.com.misvscrm.dataProviders.ConfigFileReader;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;


public class ExcelUtils {
	
	//ConfigFileReader cf = new ConfigFileReader();
	
	public Sheet createSheet(String fullURL, String newfileNameExcel){
		
		XlsxVsXls xlsx2xls = new XlsxVsXls();
		FileUtils fileUtils = new FileUtils();
    	File fileExcel = null;
		try {
			fileExcel = new File(xlsx2xls.convertXlsx2Xls(fileUtils.findFile(fullURL, newfileNameExcel)));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
    	Workbook workbookExcel = null;
		try {
			workbookExcel = Workbook.getWorkbook(fileExcel);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		//Workbook workbookExcel = createWorkBook(fullURL, newfileNameExcel);
    	Sheet sheetExcel = workbookExcel.getSheet(0);
		return sheetExcel;
	
	}
	
//	public Sheet createSheet(String fullURL, String newfileNameExcel){
//		
//		Xlsx2xls xlsx2xls = new Xlsx2xls();
//		FileUtils fileUtils = new FileUtils();
//    	File fileExcel = null;
//		try {
//			fileExcel = new File(xlsx2xls.convertXlsx2Xls(fileUtils.findFile(fullURL, newfileNameExcel)));
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//    	Workbook workbookExcel = null;
//		try {
//			workbookExcel = new Workbook();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		workbookExcel.loadFromFile(fileExcel.getAbsolutePath());
//		Worksheet worksheet = workbookExcel.getWorksheets().get(0);
//		return worksheet;
//	
//	}
	
//	public Workbook createWorkBook(String fullURL, String newfileNameExcel){
//		
//		Xlsx2xls xlsx2xls = new Xlsx2xls();
//		FileUtils fileUtils = new FileUtils();
//    	File fileExcel = null;
//		try {
//			fileExcel = new File(xlsx2xls.convertXlsx2Xls(fileUtils.findFile(fullURL, newfileNameExcel)));
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//    	Workbook workbookExcel = null;
//		try {
//			workbookExcel = Workbook.getWorkbook(fileExcel);
//		} catch (BiffException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		return workbookExcel;
//	
//	}
	
	//public String findTextInExcel(Sheet sheetMIS, String text) {
	public String findTextInExcel(Worksheet sheetMIS, String text) {

		CellRange cell = null;

		//for (int j = 1; j <= sheetMIS.getLastRow(); j++) {
			
				for (int i = 1; i <= sheetMIS.getLastColumn(); i++) {
					//System.out.println("here");
					cell = sheetMIS.getCellRange(sheetMIS.getLastRow(), i);
					if(cell.getValue().contains(text)) {				
						break;
					}

				}

		//}
		//System.out.println("aqui " + cell.getContents());
		return "aqui " + cell.getValue();

	}
	
	public int getDigitsFromString(String text) {

		int digits = Integer.valueOf(text.replaceAll("\\D+",""));
		//System.out.println("aqui2 " + digits);
		return digits;

	}
	
//	public int countExcelColumns(Sheet sheet) {
//		
//		int sheetColumnIndex = sheet.getColumns();
//		return sheetColumnIndex;
//		
//	}
//	
//	public int countExcelRows(Sheet sheet) {
//		
//		int sheetRowIndex = sheet.getRows();
//		return sheetRowIndex;
//		
//	}


}
