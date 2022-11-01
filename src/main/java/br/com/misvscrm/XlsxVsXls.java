package br.com.misvscrm;

import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

import com.aspose.cells.SaveOptions;
import  com.aspose.cells.Workbook; 

public class XlsxVsXls {
	
	
	public String convertXlsx2Xls(String path) throws Exception {
		
//		Path path2 = Paths.get(path);	
//		System.out.println(path2.getFileName());
		Workbook convertXLSX = new Workbook(path);
		//path = path.replace(".xlsx", ".xls");
		convertXLSX.save(path.replace(".xlsx", ".xls"));
		//System.out.println("aqui" + convertXLSX.getFileName());
		return convertXLSX.getFileName().toString();
    	
    	
//		Workbook convertXLSX = new Workbook(path);
//		//path = path.replace(".xlsx", ".xls");
//		convertXLSX.save(path.replace(".xlsx", ".xls"));
//    	File fileToConvert = new File(path);
//    	fileToConvert.delete();
//    	File fileConverted = new File(path.replace(".xlsx", ".xls"));
//		//System.out.println(convertXLSX.getFileName());
//		//return convertXLSX.getFileName().toString();
//    	return fileConverted;
		
	}

	public String convertXls2Xlsx(String path){
		
		Workbook convertXLS = null;
		try {
			convertXLS = new Workbook(path);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		path = path.concat("x");
		try {
			convertXLS.save(path);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		//System.out.println(convertXLSX.getFileName());
		return convertXLS.getFileName().toString();
		
//		Workbook convertXLSX = new Workbook(path);
//		//path = path.replace(".xlsx", ".xls");
//		convertXLSX.save(path.replace(".xls", ".xlsx"));
//    	File fileToConvert = new File(path);
//    	fileToConvert.delete();
//    	File fileConverted = new File(path.replace(".xls", ".xlsx"));
//		//System.out.println(convertXLSX.getFileName());
//		//return convertXLSX.getFileName().toString();
//    	return fileConverted;
		
	}
	
//	public String changeFileName(String path, String newName) {
//		
//		Path pathFile = Paths.get(path);
//		
//		return path.replace(pathFile.getFileName(), newName);
//		
//		
//		
//	}
}
