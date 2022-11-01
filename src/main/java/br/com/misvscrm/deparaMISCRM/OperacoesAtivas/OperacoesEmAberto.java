package br.com.misvscrm.deparaMISCRM.OperacoesAtivas;

import java.awt.Toolkit;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import javax.swing.ImageIcon;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import com.aspose.cells.Font;
import com.spire.xls.CellRange;
import com.spire.xls.Worksheet;
import com.spire.xls.core.converter.spreadsheet.UserBViewCollection;

import br.com.misvscrm.App;
import br.com.misvscrm.XlsxVsXls;
import br.com.misvscrm.dataProviders.ConfigFileReader;
import br.com.misvscrm.managers.XMLReaderManager;
import br.com.misvscrm.utils.ExcelUtils;
import br.com.misvscrm.utils.FileUtils;
import br.com.misvscrm.utils.InterfaceView;
import br.com.misvscrm.utils.OrderExcelUtils;
import jxl.Cell;
import jxl.Sheet;

public class OperacoesEmAberto {
	
	//static ConfigFileReader cf = new ConfigFileReader();
	//static Properties prop = null;
	static XMLReaderManager xmlReaderManager = new XMLReaderManager();
	static String typeExecutionEclipse = "eclipse";
	static String typeExecutionFolderUser = "folderUser";
	static String userDir = System.getProperty("user.dir");
	static String urlBaseEclipse = userDir + "\\src\\planilhas\\operacoes em aberto";
	static String urlBase = null;
	static String fileNameCRM = xmlReaderManager.getNodeElementText("OperaçõesAtivas", "fileNameCRM");
	static String fileNameMIS = xmlReaderManager.getNodeElementText("OperaçõesAtivas", "fileNameMIS");
	static String fileRenameCRM = xmlReaderManager.getNodeElementText("OperaçõesAtivas", "fileRenameCRM");
	static String fileRenameMIS = xmlReaderManager.getNodeElementText("OperaçõesAtivas", "fileRenameMIS");
	static String outputFilePath = null;
	public static void main(String[] args) {
		//deparaMISCRM(typeExecutionEclipse);
		//deparaMISCRM(typeExecutionFolderUser);
		
		//defineExecutionType(typeExecutionEclipse);
		defineExecutionType(typeExecutionFolderUser);
		
		outputFilePath = urlBase + xmlReaderManager.getNodeElementText("OperaçõesAtivas", "outputFile");

		InterfaceView interfaceView = new InterfaceView("CRM vs. MIS - Operações ativas");
		renameFiles();
		interfaceView.setProgressBarValue(25);
		orderFiles();
		interfaceView.setProgressBarValue(50);
		createSheetOutput2();
		interfaceView.setProgressBarValue(100);
		interfaceView.closeFrame();
		interfaceView.showMessage("Depara CRM vs. MIS para operações ativas finalizado com sucesso!");
		interfaceView.showMessage("Foi criado a planilha 'DeparaCRMvsMIS-Output.xlsx' na mesma pasta.");
		interfaceView.systemExit();
		//System.exit(0);
		
	}
	
	public static void deparaMISCRM(String typeExecution) {
	
		defineExecutionType(typeExecution);
		
		outputFilePath = urlBase + xmlReaderManager.getNodeElementText("OperaçõesAtivas", "outputFile");

		InterfaceView interfaceView = new InterfaceView("CRM vs. MIS - Operações ativas");
		renameFiles();
		interfaceView.setProgressBarValue(25);
		orderFiles();
		interfaceView.setProgressBarValue(50);
		createSheetOutput2();
		interfaceView.setProgressBarValue(100);
		interfaceView.closeFrame();
		interfaceView.showMessage("Depara CRM vs. MIS para operações ativas finalizado com sucesso!");
		interfaceView.showMessage("Foi criado a planilha 'DeparaCRMvsMIS-Output.xlsx' na mesma pasta.");
		interfaceView.systemExit();
		//System.exit(0);
		
	}
	
//	public static void renameFiles(){
//
//		FileUtils fileUtils = new FileUtils();
//
//		//rename CRM excel file
//		try {
//			fileUtils.renameFile(fileUtils.findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM()), "CRM.xlsx");
//		} catch (NullPointerException e) {
//			try {
//				fileUtils.renameFile(fileUtils.findFile(cf.getUrlOperacoesEmAberto(), "CRM.xlsx"), "CRM.xlsx");
//			} catch (Exception e1) {
//				// TODO Auto-generated catch block
//				e1.printStackTrace();
//			}
//			// TODO Auto-generated catch block
//			//e.printStackTrace();
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
// 	
//    	//rename MIS excel file
//    	try {
//			fileUtils.renameFile(fileUtils.findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoMIS()), "MIS.xlsx");
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//    	
//	}
	
	public static String defineExecutionType(String typeExecution) {
		
		if(typeExecution.equals(typeExecutionEclipse)) {
			urlBase = urlBaseEclipse;
		}
		else {
			try {
				urlBase = new File(OperacoesAtivasFolder.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent();
			} 
			catch (URISyntaxException e) {
				e.printStackTrace();
			} 
		}
		
		return urlBase;
	}
	
	public static void renameFiles(){

		FileUtils fileUtils = new FileUtils();
		
		//rename CRM excel file
		try {
			fileUtils.renameFile(fileUtils.findFile(urlBase, fileNameCRM), fileRenameCRM);
		} 
		catch (NullPointerException e) {
			try {
				fileUtils.renameFile(fileUtils.findFile(urlBase, fileRenameCRM), fileRenameCRM);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				//e1.printStackTrace();
			}
			// TODO Auto-generated catch block
			//e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			//e.printStackTrace();
		}
		
//		try {
//			fileUtils.renameFile(fileUtils.findFile(urlBaseEclipse, fileNameCRM), fileRenameCRM);
//		} catch (NullPointerException e) {
//			System.out.println("dasfdsa");
//			try {
//				fileUtils.renameFile(fileUtils.findFile(urlBaseEclipse, fileRenameCRM), fileRenameCRM);
//			} catch (Exception e1) {
//				// TODO Auto-generated catch block
//				//e1.printStackTrace();
//			}
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			//e.printStackTrace();
//		}
 	
    	//rename MIS excel file
    	try {
			fileUtils.renameFile(fileUtils.findFile(urlBase, fileNameMIS), fileRenameMIS);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			//e.printStackTrace();
		}
//    	try {
//			fileUtils.renameFile(fileUtils.findFile(urlBaseEclipse, fileNameMIS), fileRenameMIS);
//		} catch (Exception e) {
//			// TODO Auto-generated catch block
//			//e.printStackTrace();
//		}

	}
	
//	public static Sheet createSheet(String sheetName){
//		
//		ExcelUtils excelUtils = new ExcelUtils();		
//		Sheet sheet = excelUtils.createSheet(cf.getUrlOperacoesEmAberto(), sheetName);	
//		return sheet;
//	
//	}
	
	public static Worksheet createSheet(String sheetName){
		
		OrderExcelUtils orderExcelUtils = new OrderExcelUtils();	
		//Worksheet worksheet = orderExcelUtils.createSheet(cf.getUrlOperacoesEmAberto(), sheetName);
		Worksheet worksheet = orderExcelUtils.createSheet(urlBase, sheetName);
		return worksheet;
	
	}
	
	//public static void orderFiles(Worksheet sheetCRM, Worksheet sheetMIS, String excelCRM, String excelMIS) {
	//public static void orderFiles(String excelCRM, String excelMIS) {
	public static void orderFiles() {
		//String urlBase = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + "Planilhas" + File.separator + "Operações Ativas";
		
	    Worksheet sheetMIS = createSheet(fileRenameMIS);
	    Worksheet sheetCRM = createSheet(fileRenameCRM);
	    
	    FileUtils fileUtils = new FileUtils();
	    ExcelUtils excelUtils = new ExcelUtils();
		OrderExcelUtils orderExcelUtils = new OrderExcelUtils();
//		orderExcelUtils.orderExcelByColumn(9, "A1", "J" + sheetCRM.getLastRow(), fileUtils.findExactFile(cf.getUrlOperacoesEmAberto(), "CRM.xls"));
//		orderExcelUtils.orderExcelByColumn(26, "A22", "AE" + (22 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:"))), fileUtils.findExactFile(cf.getUrlOperacoesEmAberto(), "MIS.xls"));
//		orderExcelUtils.orderExcelByColumn(9, "A1", "J" + sheetCRM.getLastRow(), fileUtils.findExactFile(cf.getUrlOperacoesEmAberto(), excelCRM));
//		orderExcelUtils.orderExcelByColumn(26, "A22", "AE" + (22 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:"))), fileUtils.findExactFile(cf.getUrlOperacoesEmAberto(), excelMIS));
		orderExcelUtils.orderExcelByColumn(9, "A1", "J" + sheetCRM.getLastRow(), fileUtils.findExactFile(urlBase, fileRenameCRM));
		orderExcelUtils.orderExcelByColumn(26, "A22", "AE" + (22 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:"))), fileUtils.findExactFile(urlBase, fileRenameMIS));
		
		
	}
	
//	public static void createSheetOutput(Sheet sheetMIS, Sheet sheetCRM){
//		
//		//ExcelUtils excelUtils = new ExcelUtils();
//			
//	   	HSSFWorkbook workbook = new HSSFWorkbook();
//		 
//	   	HSSFSheet sheet = workbook.createSheet("CRM vs MIS");
//	   	 
//	   	HSSFRow rowhead;
//	   	 
//		rowhead = sheet.createRow((short)0);
//		rowhead.createCell(1).setCellValue("MIS");  
//	    rowhead.createCell(2).setCellValue("CRM");
//
//        CellStyle style = workbook.createCellStyle();
//        HSSFFont font = workbook.createFont();
//        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
//        style.setFont(font);
//	  
////	    List<String> column1 = new ArrayList<String>();
////	    column1 = new ArrayList<String>();
//	    //for (int i = 21; i < sheetMIS.getRows() - 1; i++) {
//        System.out.println(sheetCRM.getRows());
//	    for (int i = 21; i < 100; i++) {
//	    //for (int i = 21; i <= 21 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:")); i++) {
//	    	
//	        Cell cellContratoMIS = sheetMIS.getCell(26, i);
//	        //column1.add(cellContratoMIS.getContents());
//	        Cell cellValorMIS = sheetMIS.getCell(14, i);
//	        //column1.add(cellValorMIS.getContents());
//	        
//	        Cell cellCodBoletagemCRM = null;
//	        Cell cellValorCRM = null;
//	        int j;
//	        for(j = 1; j < sheetCRM.getRows(); j++) {
//	        	
//	        	cellCodBoletagemCRM = sheetCRM.getCell(9, j);
//	            //column1.add(cellCodBoletagemCRM.getContents());
//	            cellValorCRM = sheetCRM.getCell(7, j);
//	            //column1.add(cellValorCRM.getContents());        	   
//	
//	            if(cellContratoMIS.getContents().contentEquals(cellCodBoletagemCRM.getContents())) {
//	            	System.out.println(cellContratoMIS.getContents() + "		" + cellCodBoletagemCRM.getContents());
//	            	break;
//	            }
//	            
//	        }
//	       
//	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
//	        //if(!cellValorMIS.getContents().contentEquals(cellValorCRM.getContents())) {
//	        	
//		        rowhead = sheet.createRow((short)(i-21)+1);
//		        rowhead.createCell(0).setCellValue(cellContratoMIS.getContents());  
//		        //rowhead.createCell(1).setCellValue(cellValorMIS.getContents().replace("[$-416]", "").replace(".", ","));
//		        rowhead.createCell(1).setCellValue(cellValorMIS.getContents());
//		        rowhead.createCell(2).setCellValue(cellValorCRM.getContents().toString().replace(",00", ""));
//		        //rowhead.getCell(2).setCellStyle(style);
//	        
//	        //}
//	        //rowhead.getCell(2).setCellStyle(style);
//	        if(!rowhead.getCell(1).toString().equals(rowhead.getCell(2).toString())) {
//	        	System.out.println("aqui: " + rowhead.getCell(1) + "     " + rowhead.getCell(2));
//	        	rowhead.getCell(2).setCellStyle(style);
//	        }
//	        if(i == 21) {
//	        	rowhead.createCell(2).setCellValue("Valor Original da Operação");
//	        	System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
//	        }
//	        else {
//	        	System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + cellValorCRM.getContents());
//	        }
//	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
//	    }
//	
//		FileOutputStream fileOut = null;
//		try {
//			fileOut = new FileOutputStream("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		}
//		 try {
//			workbook.write(fileOut);
//			fileOut.close();
//			workbook.close();
//		} catch (IOException e) {
//	
//			e.printStackTrace();
//		}
//		
//	}
	
	public static void createSheetOutput(Worksheet sheetMIS, Worksheet sheetCRM){
		
		//ExcelUtils excelUtils = new ExcelUtils();
	   	HSSFWorkbook workbook = new HSSFWorkbook();
		 
	   	HSSFSheet sheet = workbook.createSheet("CRM vs MIS");
	   	 
	   	HSSFRow rowhead;
	   	 
		rowhead = sheet.createRow((short)0);
		rowhead.createCell(1).setCellValue("MIS");  
	    rowhead.createCell(2).setCellValue("CRM");

        CellStyle style = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        style.setFont(font);
	  
//	    List<String> column1 = new ArrayList<String>();
//	    column1 = new ArrayList<String>();
        
        ExcelUtils excelUtils = new ExcelUtils();
	    //for (int i = 21; i < sheetMIS.getRows() - 1; i++) {
	    //for (int i = 22; i < 100; i++) {
	    for (int i = 22; i <= 22 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:")); i++) {
	    	
	        CellRange cellContratoMIS = sheetMIS.getCellRange(i, 27);
	        //column1.add(cellContratoMIS.getContents());
	        CellRange cellValorMIS = sheetMIS.getCellRange(i, 15);
	        //column1.add(cellValorMIS.getContents());
	        
	        CellRange cellCodBoletagemCRM = null;
	        CellRange cellValorCRM = null;
	        int j;
	        for(j = 1; j < sheetCRM.getLastRow() + 1; j++) {
	        	
	        	cellCodBoletagemCRM = sheetCRM.getCellRange(j, 10);
	            //column1.add(cellCodBoletagemCRM.getContents());
	            //cellValorCRM = sheetCRM.getCellRange(j, 8);
	            //column1.add(cellValorCRM.getContents());   

	
	            if(cellContratoMIS.getValue().contentEquals(cellCodBoletagemCRM.getValue())) {
	            	System.out.println(cellContratoMIS.getValue() + "		" + cellCodBoletagemCRM.getValue());
	            	cellValorCRM = sheetCRM.getCellRange(j, 8);
	            	break;
	            }
	            
	        }
	       
	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
	        //if(!cellValorMIS.getContents().contentEquals(cellValorCRM.getContents())) {
	        	
		        rowhead = sheet.createRow((short)(i-22)+1);
		        rowhead.createCell(0).setCellValue(cellContratoMIS.getValue());  
		        //rowhead.createCell(1).setCellValue(cellValorMIS.getContents().replace("[$-416]", "").replace(".", ","));
		        rowhead.createCell(1).setCellValue(cellValorMIS.getValue());
		        if(cellValorCRM != null) {
		        	rowhead.createCell(2).setCellValue(cellValorCRM.getValue());
		        }
		        else {
		        	rowhead.createCell(2).setCellValue("Inexistente");
		        }
		        //rowhead.createCell(2).setCellValue(cellValorCRM.getValue().toString().replace(",00", ""));
		        //rowhead.getCell(2).setCellStyle(style);
	        
	        //}
	        //rowhead.getCell(2).setCellStyle(style);
	        //if(!rowhead.getCell(1).toString().equals(rowhead.getCell(2).toString())) {
	        if(i != 22 && cellValorCRM != null && Float.compare(Float.valueOf(rowhead.getCell(1).toString().replace(",", ".")), Float.valueOf(rowhead.getCell(2).toString().replace(",", "."))) != 0) {
	        	System.out.println("aqui: " + rowhead.getCell(1) + "     " + rowhead.getCell(2));
	        	rowhead.getCell(2).setCellStyle(style);
	        }
	        if(i == 22) {
	        	rowhead.createCell(2).setCellValue("Valor Original da Operação");
	        	System.out.println(cellContratoMIS.getValue() + "		" + cellValorMIS.getValue().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
	        }
	        else{
	        	System.out.println(cellContratoMIS.getValue() + "		" + cellValorMIS.getValue().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
	        }
	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
	    }
	
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		 try {
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		} catch (IOException e) {
	
			e.printStackTrace();
		}
		 
		OrderExcelUtils orderExcelUtils = new OrderExcelUtils();
		System.out.println(sheet.getLastRowNum() + 1);
		orderExcelUtils.orderExcelByColumn(0, "A2", "C" + (sheet.getLastRowNum()+1), "src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
		
	}
	
	//public static void createSheetOutput2(Worksheet sheetMIS, Worksheet sheetCRM){
	public static void createSheetOutput2(){
		
	    Worksheet sheetMIS = createSheet(fileRenameMIS);
	    Worksheet sheetCRM = createSheet(fileRenameCRM);
		
		//ExcelUtils excelUtils = new ExcelUtils();
	   	HSSFWorkbook workbook = new HSSFWorkbook();
		 
	   	HSSFSheet sheet = workbook.createSheet("CRM vs MIS");
	   	 
	   	HSSFRow rowhead;
	   	 
		rowhead = sheet.createRow((short)0);
		rowhead.createCell(1).setCellValue("CRM");  
	    rowhead.createCell(2).setCellValue("MIS");

        CellStyle style = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        style.setFont(font);
	  
//	    List<String> column1 = new ArrayList<String>();
//	    column1 = new ArrayList<String>();
        
        ExcelUtils excelUtils = new ExcelUtils();
	    //for (int i = 21; i < sheetMIS.getRows() - 1; i++) {
	    //for (int i = 22; i < 100; i++) {
        CellRange cellCodBoletagemCRM = null;
        CellRange cellValorCRM = null;
        int k = 23;
        for(int j = 1; j < sheetCRM.getLastRow() + 1; j++) {
        //for(int j = 1; j < 2 + 1; j++) {
        	
	    //for (int i = 22; i <= 22 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:")); i++) {
	    	
//	        CellRange cellContratoMIS = sheetMIS.getCellRange(i, 27);
//	        //column1.add(cellContratoMIS.getContents());
//	        CellRange cellValorMIS = sheetMIS.getCellRange(i, 15);
//	        //column1.add(cellValorMIS.getContents());
	        
	        cellCodBoletagemCRM = sheetCRM.getCellRange(j, 10);
	        cellValorCRM = sheetCRM.getCellRange(j, 9);
	        

	        
	        CellRange cellContratoMIS = null, cellValorMIS = null;
	        System.out.println(j);
	        for (int i = k; i <= 23 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:")) && j != 1; i++) {
	        //for(j = 1; j < sheetCRM.getLastRow() + 1; j++) {
	        	
//	        	cellCodBoletagemCRM = sheetCRM.getCellRange(j, 10);
//	            //column1.add(cellCodBoletagemCRM.getContents());
//	            //cellValorCRM = sheetCRM.getCellRange(j, 8);
//	            //column1.add(cellValorCRM.getContents()); 
	        	
		        cellContratoMIS = sheetMIS.getCellRange(i, 27);
		        //column1.add(cellContratoMIS.getContents());
		        //cellValorMIS = sheetMIS.getCellRange(i, 15);

		        //System.out.println("confere " + cellCodBoletagemCRM.getValue() + "		" + cellContratoMIS.getValue());
	            if(cellCodBoletagemCRM.getValue().equals(cellContratoMIS.getValue())) {
	            	System.out.println(cellCodBoletagemCRM.getValue() + "		" + cellContratoMIS.getValue());
	            	cellValorMIS = sheetMIS.getCellRange(i, 15);
	            	//System.out.println(cellValorMIS.getValue());
	            	k = i + 1;
//	            	System.out.println(i);
//	            	System.out.println(k);
	            	break;
	            }
	            
	        }
	       
	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
	        //if(!cellValorMIS.getContents().contentEquals(cellValorCRM.getContents())) {
	        	
		        rowhead = sheet.createRow((short)j);
		        rowhead.createCell(0).setCellValue(cellCodBoletagemCRM.getValue());  
		        //rowhead.createCell(1).setCellValue(cellValorMIS.getContents().replace("[$-416]", "").replace(".", ","));
		        rowhead.createCell(1).setCellValue(cellValorCRM.getValue());
		        //rowhead.createCell(2).setCellValue(cellValorMIS.getValue());
		        if(cellValorMIS != null) {
		        	rowhead.createCell(2).setCellValue(cellValorMIS.getValue());
		        }
		        else {
		        	rowhead.createCell(2).setCellValue("Inexistente");
		        }
		        //rowhead.createCell(2).setCellValue(cellValorCRM.getValue().toString().replace(",00", ""));
		        //rowhead.getCell(2).setCellStyle(style);
	        
	        //}
	        //rowhead.getCell(2).setCellStyle(style);
	        //if(!rowhead.getCell(1).toString().equals(rowhead.getCell(2).toString())) {
//		    System.out.println("aqui: " + rowhead.getCell(1) + "     " + rowhead.getCell(2));
//		    if(rowhead.getCell(2).getStringCellValue().equals("Inexistente")) {
//		    	System.out.println("here");    	
//		    }
		    //!rowhead.getCell(2).getStringCellValue().equals("0") &&
	        if(j != 1 && !rowhead.getCell(2).getStringCellValue().equals("Inexistente") && !rowhead.getCell(1).getStringCellValue().isEmpty() &&
	        	Float.compare(Float.valueOf(rowhead.getCell(1).toString().replace(",", ".")), Float.valueOf(rowhead.getCell(2).toString().replace(",", "."))) != 0) {
	        	System.out.println("aqui: " + rowhead.getCell(1) + "     " + rowhead.getCell(2));
	        	rowhead.getCell(2).setCellStyle(style);
	        }
	        if(j == 1) {
	        	rowhead.createCell(2).setCellValue("Outstanding");
	        	System.out.println(cellCodBoletagemCRM.getValue() + "		" + cellValorCRM.getValue().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
	        }
	        else{
	        	System.out.println(cellCodBoletagemCRM.getValue() + "		" + cellValorCRM.getValue().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
	        }
	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
	    }
	
		FileOutputStream fileOut = null;
		try {
			//fileOut = new FileOutputStream("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
			fileOut = new FileOutputStream(outputFilePath);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		 try {
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		} catch (IOException e) {
	
			e.printStackTrace();
		}
		 
		OrderExcelUtils orderExcelUtils = new OrderExcelUtils();
		System.out.println(sheet.getLastRowNum() + 1);
		//orderExcelUtils.orderExcelByColumn(0, "A2", "C" + (sheet.getLastRowNum()+1), "src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
		orderExcelUtils.orderExcelByColumn(0, "A2", "C" + (sheet.getLastRowNum()+1), outputFilePath);
		XlsxVsXls xlsxVsXls = new XlsxVsXls();
		xlsxVsXls.convertXls2Xlsx(outputFilePath);
		FileUtils fileUtils = new FileUtils();
		fileUtils.deleteFile(outputFilePath);
		
	}
	
//	public static void createSheetOutput(Sheet sheetMIS, Sheet sheetCRM){
//		
//		//ExcelUtils excelUtils = new ExcelUtils();
//			
//	   	HSSFWorkbook workbook = new HSSFWorkbook();
//		 
//	   	HSSFSheet sheet = workbook.createSheet("CRM vs MIS");
//	   	 
//	   	HSSFRow rowhead;
//	   	 
//		rowhead = sheet.createRow((short)0);
//		rowhead.createCell(1).setCellValue("MIS");  
//	    rowhead.createCell(2).setCellValue("CRM");
//
//        CellStyle style = workbook.createCellStyle();
//        HSSFFont font = workbook.createFont();
//        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
//        style.setFont(font);
//	  
////	    List<String> column1 = new ArrayList<String>();
////	    column1 = new ArrayList<String>();
//        ExcelUtils excelUtils = new ExcelUtils();
//	    //for (int i = 21; i < sheetMIS.getRows() - 1; i++) {
//	    //for (int i = 21; i < 100; i++) {
//        Cell cellCodBoletagemCRM = null, cellValorCRM = null, cellContratoMIS = null, cellValorMIS = null;
//        //for (int i = 1; i <= sheetCRM.getRows(); i++) {
//        System.out.println("aqui");
//	    for (int i = 1; i <= 50; i++) {
//	    	System.out.println(i);
//        	cellCodBoletagemCRM = sheetCRM.getCell(9, i);
//            //column1.add(cellCodBoletagemCRM.getContents());
//            cellValorCRM = sheetCRM.getCell(7, i);
//            //column1.add(cellValorCRM.getContents()); 
//
//	        int j;
//	        for(j = 21; j <= 21 + excelUtils.getDigitsFromString(excelUtils.findTextInExcel(sheetMIS, "Total de Registros:")); j++) {
//	        	System.out.println(j);
//		        cellContratoMIS = sheetMIS.getCell(26, j);
//		        //column1.add(cellContratoMIS.getContents());
//		        cellValorMIS = sheetMIS.getCell(14, j);
//		        //column1.add(cellValorMIS.getContents());        	   
//	
//	            if(cellCodBoletagemCRM.getContents().contentEquals(cellContratoMIS.getContents())) {
//	            	System.out.println(cellCodBoletagemCRM.getContents() + "		" + cellContratoMIS.getContents());
//	            	break;
//	            }
//	            
//	        }
//	       
//	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
//	        //if(!cellValorMIS.getContents().contentEquals(cellValorCRM.getContents())) {
//	        	System.out.println(i);
//		        rowhead = sheet.createRow((short)i + 1);
//		        rowhead.createCell(0).setCellValue(cellCodBoletagemCRM.getContents());  
//		        //rowhead.createCell(1).setCellValue(cellValorMIS.getContents().replace("[$-416]", "").replace(".", ","));
////		        rowhead.createCell(1).setCellValue(cellValorCRM.getContents().toString().replace(",00", ""));
////		        rowhead.createCell(2).setCellValue(cellValorMIS.getContents());
//		        rowhead.createCell(1).setCellValue(Float.valueOf(cellValorCRM.getContents()));
//		        rowhead.createCell(2).setCellValue(Float.valueOf(cellValorMIS.getContents()));
//		        //rowhead.getCell(2).setCellStyle(style);
//	        
//	        //}
//	        //rowhead.getCell(2).setCellStyle(style);
//	        if(!rowhead.getCell(1).toString().equals(rowhead.getCell(2).toString())) {
//	        	System.out.println("aqui: " + rowhead.getCell(1) + "     " + rowhead.getCell(2));
//	        	rowhead.getCell(2).setCellStyle(style);
//	        }
//	        if(i == 21) {
//	        	rowhead.createCell(2).setCellValue("Valor Original da Operação");
//	        	System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
//	        }
//	        else {
//	        	System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + cellValorCRM.getContents());
//	        }
//	        //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
//	    }
//	
//		FileOutputStream fileOut = null;
//		try {
//			fileOut = new FileOutputStream("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		}
//		 try {
//			workbook.write(fileOut);
//			fileOut.close();
//			workbook.close();
//		} catch (IOException e) {
//	
//			e.printStackTrace();
//		}
//		
//	}

}
