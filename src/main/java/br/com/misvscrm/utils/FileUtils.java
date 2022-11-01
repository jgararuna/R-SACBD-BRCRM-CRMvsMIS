package br.com.misvscrm.utils;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import br.com.misvscrm.deparaMISCRM.OperacoesAtivas.OperacoesAtivasFolder;
import br.com.misvscrm.managers.XMLReaderManager;

public class FileUtils {

    public void renameFile(String fileToMove, String targetFile) throws Exception {
    	
//    	ConfigFileReader cf = new ConfigFileReader();
//    	Xlsx2xls xlsx2xls = new Xlsx2xls();
//    	Path fileToMovePath = Paths.get(fileToMove);
//    	Path targetPath = Paths.get(targetFile);
//    	Files.move(fileToMovePath, targetPath);
//    	File targetFileDelete = new File(fileToMovePath.toString());
//    	targetFileDelete.delete();
    	System.out.println(fileToMove);
    	
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
    
    public String findFile(String path, String containedString) {
    	
        File[] files = new File(path).listFiles();
        String s = null;
        
        for(File f : files){
            //if(f.getName().toLowerCase().indexOf(containedString.toLowerCase()) != -1)
            if(f.getName().toLowerCase().contains(containedString.toLowerCase()))
            s = f.getPath().replace("\\", "/");
            //System.out.println(s);
        }
        
        return s;
    }
    
    public String findExactFile(String path, String fileName) {
    	
        File[] files = new File(path).listFiles(); //X is the directory
        String s = null;
        for(File f : files){
            //if(f.getName().toLowerCase().indexOf(containedString.toLowerCase()) != -1)
            if(f.getName().toLowerCase().equals(fileName.toLowerCase()))
            s = f.getPath().replace("\\", "/");
            //System.out.println(s);
        }
        
        return s;
    }
    
    public void createFolder(String folderName) {
    	
    	Path path = Paths.get(System.getProperty("user.home") + File.separator + "Desktop" + File.separator + folderName);
    	if (Files.notExists(path)) {
    		System.out.println("aqui");
    		new File(System.getProperty("user.home") + File.separator + "Desktop" + File.separator + folderName).mkdirs();
    	}
    }
    
    public String getFolderUrl(String executionType) {
    	
    	String userDir = null;
		try {
			userDir = new File(OperacoesAtivasFolder.class.getProtectionDomain().getCodeSource().getLocation().toURI()).getParent();
		} catch (URISyntaxException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return userDir;
    }
    
    public void deleteFile(String path) {
    	
		File file = new File(path);
		file.delete();
		
    }
}
