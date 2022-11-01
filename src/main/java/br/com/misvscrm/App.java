package br.com.misvscrm;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.aspose.cells.ReplaceOptions;

import br.com.misvscrm.dataProviders.ConfigFileReader;

import org.apache.commons.math3.util.ArithmeticUtils;
import static java.nio.file.StandardCopyOption.*;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {
        //System.out.println( "Hello World!" );
    	
    	ConfigFileReader cf = new ConfigFileReader();
    	//System.out.println(cf.getUrlOperacoesEmAberto());
    	//ClassLoader classLoader = ClassLoader.class.getClassLoader();
    	// "C:\\Users\\joao.araruna\\Downloads\\Oportunidades Abertas 14-09-2022 18-05-02.xlsx"
    	//File file = new File("C:\\Users\\joao.araruna\\Downloads\\Oportunidades Abertas 14-09-2022 18-05-02.xls");
    	XlsxVsXls xlsx2xls = new XlsxVsXls();
    	//File file = new File(xlsx2xls.convertXlsx2Xls("C:\\Users\\joao.araruna\\Downloads\\Oportunidades Abertas 14-09-2022 18-05-02.xlsx"));
    	//File file = new File(xlsx2xls.convertXlsx2Xls("/misvscrm/src/main/java/Oportunidades Abertas 14-09-2022 18-05-02.xlsx"));
    	//File file = new File("src/planilhas/operacoes em aberto/Oportunidades Abertas 14-09-2022 18-05-02.xlsx");
    	//new App().findFile("src/planilhas/operacoes em aberto/", "Oportunidades Abertas");
    	//File file = new File(xlsx2xls.convertXlsx2Xls(cf.getUrlOperacoesEmAberto()));
    	//File file = new File(xlsx2xls.convertXlsx2Xls(new App().findFile("src/planilhas/operacoes em aberto/", "Oportunidades Abertas")));
    	
//    	new App().renameFile(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM()), "src/planilhas/operacoes em aberto/CRM.xls");
//    	new App().renameFile(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoMIS()), "src/planilhas/operacoes em aberto/MIS.xls");
    	
//    	Path fileToMovePathCRM = Paths.get(xlsx2xls.convertXlsx2Xls(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM())));
//    	Path targetPathCRM = Paths.get("src/planilhas/operacoes em aberto/CRM.xls");
//    	Files.move(fileToMovePathCRM, targetPathCRM);
//    	File targetFileCRM = new File(fileToMovePathCRM.toString());
//    	targetFileCRM.delete();
    	
//    	File fileCRM = new File(xlsx2xls.convertXlsx2Xls(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM())));
//    	//System.out.println(file.getPath());
//    	File newFileCRM = new File("src/planilhas/operacoes em aberto/CRM.xls");
//    	fileCRM.renameTo(newFileCRM);
//    	Workbook workbookCRM =  Workbook.getWorkbook(newFileCRM);
//    	Sheet sheetCRM = workbookCRM.getSheet(0);
    	
//    	Path fileToMovePath = Paths.get(xlsx2xls.convertXlsx2Xls(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoMIS())));
//    	Path targetPath = Paths.get("src/planilhas/operacoes em aberto/MIS.xls");
//    	Files.move(fileToMovePath, targetPath);
//    	File targetFile = new File(fileToMovePath.toString());
//    	targetFile.delete();
    	
    	App app = new App();
    	
        //Path oldFile = Paths.get("src/planilhas/operacoes em aberto/MIS5500.xlsx");
        //System.out.println(new App().findFile(cf.getUrlOperacoesEmAberto(), "MIS"));
    	try {
    		app.renameFile(app.findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM()), "CRM.xlsx");
		} catch (NullPointerException e) {
			app.renameFile(app.findFile(cf.getUrlOperacoesEmAberto(), "CRM.xlsx"), "CRM.xlsx");
		}
    	//app.renameFile(app.findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM()), "CRM.xlsx");
    	//File fileCRM = new File(app.findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoCRM()), "CRM.xls");
    	File fileCRM = new File(xlsx2xls.convertXlsx2Xls(app.findFile(cf.getUrlOperacoesEmAberto(), "CRM.xlsx")));
    	Workbook workbookCRM =  Workbook.getWorkbook(fileCRM);
    	Sheet sheetCRM = workbookCRM.getSheet(0);
    	app.renameFile(app.findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoMIS()), "MIS.xlsx");
    	File fileMIS = new File(xlsx2xls.convertXlsx2Xls(app.findFile(cf.getUrlOperacoesEmAberto(), "MIS.xlsx")));
    	Workbook workbookMIS =  Workbook.getWorkbook(fileMIS);
    	Sheet sheetMIS = workbookMIS.getSheet(0);
    	
    	
        
//        Path oldFile = Paths.get(new App().findFile(cf.getUrlOperacoesEmAberto(), "MIS"));
//
//	    try {
//	        Files.move(oldFile, oldFile.resolveSibling(
//	                                "MIS.xls"));
//	        System.out.println("File Successfully Rename");
//	    }
//	    catch (IOException e) {
//	        System.out.println("operation failed");
//	    }
    	
//    	File fileMIS = new File(xlsx2xls.convertXlsx2Xls(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoMIS())));
//    	File newFileMIS = new File("src/planilhas/operacoes em aberto/MIS.xls");
//    	fileMIS.renameTo(newFileMIS);
//    	Workbook workbookMIS =  Workbook.getWorkbook(newFileMIS);
//    	Sheet sheetMIS = workbookMIS.getSheet(0);
  	 
//    	 int column_index_1 = 0;
//    	 int column_index_2 = 0;
//    	 int column_index_3 = 0;
//    	 Row row = sheet.getRow(0);
    	 
//    	 for (Cell cell : row) {
//    	        // Column header names.
//    	        switch (cell.) {
//    	            case "MyFirst Column":
//    	                column_index_1 = cell.getColumnIndex();
//    	                break;
//    	            case "3rd Column":
//    	                column_index_2 = cell.getColumnIndex();
//    	                break;
//    	            case "forth Column":
//    	                column_index_3 = cell.getColumnIndex();
//    	                break;
//    	        }
//    	    }
    	
    	 int masterSheetColumnIndexCRM = sheetCRM.getColumns();
    	 //System.out.println("Número de colunas: " + masterSheetColumnIndexCRM);
    	 int masterSheetRowIndexCRM = sheetCRM.getColumn(0).length;
    	 //System.out.println("Número de registros: " + masterSheetRowIndexCRM);
    	 
    	 int masterSheetColumnIndexMIS = sheetMIS.getColumns();
    	 //System.out.println("Número de colunas: " + masterSheetColumnIndexMIS);
    	 int masterSheetRowIndexMIS = sheetMIS.getColumn(0).length;
    	 //System.out.println("Número de registros: " + masterSheetRowIndexMIS);
//    	 List<String> ExpectedColumns = new ArrayList<String>();
//    	    for (int x = 0; x < masterSheetColumnIndexMIS; x++) {
//    	        Cell celll = sheetMIS.getCell(x, 0);
//    	        String d = celll.getContents();
//    	        ExpectedColumns.add(d);
//    	        System.out.println(d);
//    	 }
    	 
    	 HSSFWorkbook workbook2 = new HSSFWorkbook();
    	 
    	 HSSFSheet sheet2 = workbook2.createSheet("CRM vs MIS");
    	 
    	 HSSFRow rowhead;
    	 
    	 rowhead = sheet2.createRow((short)0);
    	 rowhead.createCell(1).setCellValue("MIS");  
         rowhead.createCell(2).setCellValue("CRM");  
    	 
    	 //FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\joao.araruna\\Downloads\\Output.xlsx"));
    	 
    	 LinkedHashMap<String, List<String>> columnDataValues = new LinkedHashMap<String, List<String>>();
    	 List<String> column1 = new ArrayList<String>();
    	 //for (int j = 0; j < 1; j++) {
    	        column1 = new ArrayList<String>();
    	        //for (int i = 1; i < sheet.getRows(); i++) {
    	        for (int i = 21; i < masterSheetRowIndexMIS - 1; i++) {
    	        	
    	            Cell cellContratoMIS = sheetMIS.getCell(26, i);
    	            column1.add(cellContratoMIS.getContents());
    	            Cell cellValorMIS = sheetMIS.getCell(14, i);
    	            column1.add(cellValorMIS.getContents());
    	            
    	            Cell cellCodBoletagemCRM = null;
    	            Cell cellValorCRM = null;
    	            int j;
    	            for(j = 1; j < masterSheetRowIndexCRM; j++) {
    	            	
    	            	cellCodBoletagemCRM = sheetCRM.getCell(9, j);
        	            column1.add(cellCodBoletagemCRM.getContents());
        	            cellValorCRM = sheetCRM.getCell(7, j);
        	            column1.add(cellValorCRM.getContents());        	   

        	            if(cellContratoMIS.getContents().contentEquals(cellCodBoletagemCRM.getContents())) {
        	            	System.out.println(cellContratoMIS.getContents() + "		" + cellCodBoletagemCRM.getContents());
        	            	break;
        	            }
        	            
    	            }
    	           
    	            //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1)); 
    	            rowhead = sheet2.createRow((short)(i-21)+1);
    	            rowhead.createCell(0).setCellValue(cellContratoMIS.getContents());  
    	            rowhead.createCell(1).setCellValue(cellValorMIS.getContents().replace("[$-416]", "").replace(".", ""));
    	            rowhead.createCell(2).setCellValue(cellValorCRM.getContents());
    	            //System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + cellValorCRM.getContents());
    	            if(i == 21) {
    	            	rowhead.createCell(2).setCellValue("Valor Estimado da Operação");
    	            	System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + rowhead.getCell(2));
    	            }
    	            else {
    	            	System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", "") + "		" + cellValorCRM.getContents());
    	            }
    	            //System.out.println(cellContratoMIS.getContents() + "		" + cellValorMIS.getContents().replace("[$-416]", "").replace(".", ""));
    	            //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
    	        }
    	        //System.out.println(column1);
    	        //columnDataValues.put(ExpectedColumns.get(j), column1);
    	    //}
    	
//    	 List<String> column2 = new ArrayList<String>();
//    	 for (int j = 0; j < 1; j++) {
//    	        column1 = new ArrayList<String>();
//    	        //for (int i = 1; i < sheet.getRows(); i++) {
//    	        for (int i = 0; i < 40; i++) {
//    	        	
//    	            Cell cellTopico = sheetCRM.getCell(9, i);
//    	            column1.add(cellTopico.getContents());
//    	            Cell cellValor = sheetCRM.getCell(7, i);
//    	            column1.add(cellValor.getContents());
//    	            //System.out.println(column1.get(2*i) + "		" + column1.get(2*i + 1));
//    	            
//    	            rowhead = sheet2.createRow((short)i+1);
//    	            rowhead.createCell(0).setCellValue(cellTopico.getContents());  
//    	            rowhead.createCell(1).setCellValue(cellValor.getContents());
//    	            if(i == 0) {
//    	            	rowhead.createCell(2).setCellValue("Valor Estimado da Operação");
//    	            }
//    	           
//    	            System.out.println(cellTopico.getContents() + "		" + cellValor.getContents());
//    	            
//    	        }
//    	        //System.out.println(column1);
//    	        //columnDataValues.put(ExpectedColumns.get(j), column1);
//    	    }
    	 
    	 //FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\joao.araruna\\Downloads\\" + file.getName().toString().replace(".xls", "") + "Output.xls"));
    	 //FileOutputStream fileOut = new FileOutputStream(new File(cf.getUrlOperacoesEmAberto().replace(".xlsx", "") + "Output.xls"));
    	 FileOutputStream fileOut = new FileOutputStream("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
    	 workbook2.write(fileOut);
    	 //xlsx2xls.convertXls2Xlsx("C:\\Users\\joao.araruna\\Downloads\\" + file.getName().toString().replace(".xls", "") + "Output.xls");
    	 fileOut.close();
    	 workbook2.close();  
    	 
    }
    
    //new App().renameFile(new App().findFile(cf.getUrlOperacoesEmAberto(), cf.getfileNameOperacoesEmAbertoMIS()), "src/planilhas/operacoes em aberto/MIS.xls");
    
    public String findFile(String path, String containedString) {
    	
        File[] files = new File(path).listFiles(); //X is the directory
        String s = null;
        for(File f : files){
            //if(f.getName().toLowerCase().indexOf(containedString.toLowerCase()) != -1)
            if(f.getName().toLowerCase().contains(containedString.toLowerCase()))
            s = f.getPath().replace("\\", "/");
            //System.out.println(s);
        }
        
        return s;
    }
    
    public void renameFile(String fileToMove, String targetFile) throws Exception {
    	
    	ConfigFileReader cf = new ConfigFileReader();
//    	Xlsx2xls xlsx2xls = new Xlsx2xls();
//    	Path fileToMovePath = Paths.get(fileToMove);
//    	Path targetPath = Paths.get(targetFile);
//    	Files.move(fileToMovePath, targetPath);
//    	File targetFileDelete = new File(fileToMovePath.toString());
//    	targetFileDelete.delete();
    	
        Path oldFile = Paths.get(fileToMove);

	    try {
	        Files.move(oldFile, oldFile.resolveSibling(
	                                targetFile));
	        System.out.println("File Successfully Rename");
	    }
	    catch (IOException e) {
	        System.out.println("operation failed");
	    }
    	
    }
}
