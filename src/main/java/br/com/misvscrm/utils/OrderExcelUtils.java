package br.com.misvscrm.utils;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import com.spire.xls.ExcelVersion;
import com.spire.xls.OrderBy;
import com.spire.xls.SortComparsionType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import br.com.misvscrm.XlsxVsXls;
import jxl.Sheet;
import jxl.read.biff.BiffException;

public class OrderExcelUtils {
	
	public void orderExcelByColumn(int columnPosition, String firstCell, String lastCell, String pathFile) {
		
		Workbook workbook = new Workbook();
		
		workbook.loadFromFile(pathFile);
		
		Worksheet worksheet = workbook.getWorksheets().get(0);
		
		workbook.getDataSorter().getSortColumns().add(columnPosition, SortComparsionType.Values, OrderBy.Ascending);
		
		//workbook.getDataSorter().ge
		
		workbook.getDataSorter().sort(worksheet.getCellRange(firstCell + ":" + lastCell));
		
		//workbook.getDataSorter().getSortColumns().add(0, OrderBy.Ascending);
		
		workbook.saveToFile(pathFile, ExcelVersion.Version2013);		
		
	}
	
	public Worksheet createSheet(String fullURL, String newfileNameExcel){
		
		//XlsxVsXls xlsx2xls = new XlsxVsXls();
		FileUtils fileUtils = new FileUtils();
    	File fileExcel = null;
		try {
			//fileExcel = new File(xlsx2xls.convertXlsx2Xls(fileUtils.findFile(fullURL, newfileNameExcel)));
			fileExcel = new File(fileUtils.findFile(fullURL, newfileNameExcel));
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
//    	Workbook workbookExcel = null;
//		try {
//			workbookExcel = workbookExcel.loadFromFile(fileExcel.getPath());
//		} catch (BiffException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
		
		Workbook workbook = new Workbook();
		
		workbook.loadFromFile(fileExcel.getPath());
		
		Worksheet worksheet = workbook.getWorksheets().get(0);

		return worksheet;
	
	}

}
