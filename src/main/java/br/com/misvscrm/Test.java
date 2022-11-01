package br.com.misvscrm;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Panel;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.swing.ImageIcon;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JTextArea;
import javax.swing.Timer;

import com.spire.xls.ExcelVersion;
import com.spire.xls.OrderBy;
import com.spire.xls.SortComparsionType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

public class Test {
	
	
	public static void main(String[] args) {
		
		Test test = new Test();
		//name();
		test.name();
		
	}
	
	public void name() {
      
//		Workbook workbook = new Workbook();
//		
//		workbook.loadFromFile("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls");
//		
//		Worksheet worksheet = workbook.getWorksheets().get(0);
//		
//		workbook.getDataSorter().getSortColumns().add(0, SortComparsionType.Values, OrderBy.Ascending);
//		
//		//workbook.getDataSorter().ge
//		
//		workbook.getDataSorter().sort(worksheet.getCellRange("A2:C80"));
//		
//		//workbook.getDataSorter().getSortColumns().add(0, OrderBy.Ascending);
//		
//		workbook.saveToFile("src/planilhas/operacoes em aberto/CRMvsMIS-Output.xls", ExcelVersion.Version2013);
		
//        JFrame frame = new JFrame("HelloWorldSwing");
//        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//
//        JLabel label = new JLabel("Hello World");
//        frame.add(label);
//
//        //Display the window.
//        frame.pack();
//        frame.setVisible(true);
		
//		ImageIcon icon;
////		//URL img = this.getClass().getResource("loading.gif");
//		URL url = null;
//		try {
//			url = Paths.get("target", "loading.gif").toUri().toURL();
//		} catch (MalformedURLException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
////		System.out.println(url);
////		//System.out.println(this.getClass().getResource(url));
//		icon = new ImageIcon(url);
		
//		JFrame frame = new JFrame("CRM vs. MIS - Operações ativas");
////		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////		JOptionPane pane = new JOptionPane("Carregando depara das operações ativas...", JOptionPane.INFORMATION_MESSAGE, JOptionPane.DEFAULT_OPTION, null, new Object[]{}, null);
////		pane.setIcon(icon);
////		pane.setVisible(true);	
//		//JOptionPane pane = new JOptionPane("Carregando depara das operações ativas...", JOptionPane.INFORMATION_MESSAGE, JOptionPane.DEFAULT_OPTION, null, new Object[]{}, null);
////		JOptionPane pane = new JOptionPane();
////		JOptionPane.showMessageDialog(frame, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "CRM vs. MIS - Operações ativas", 0, null);
//		JProgressBar jb = new JProgressBar(1, 100);
//		//jb.setValue(15);
//		jb.setBounds(40,40,160,30);   
////		pane.add(jProgressBar,1);
////		pane.setVisible(true);
//		jb.setValue(0);    
//		jb.setStringPainted(true);  
//		frame.add(jb);
//		frame.setSize((frame.getTitle().length() * 12),150);    
//		frame.setLayout(null);
//		frame.setLocationRelativeTo(null);
//		frame.setVisible(true);
//		int i = 0;
//		while(i<=10) {
//			
//			try {
//				Thread.sleep(200);
//			} catch (InterruptedException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//			i++;
//		}
//		jb.setValue(26);
//		jb.setVisible(false);
//		
//		//label1.setVisible(true);
//		System.out.println("aqui");
////		frame.getContentPane().add(new JTextArea("Text Area\nText Area", 18,18));
//
//		//frame.set("Depara CRM vs. MIS para operações ativas finalizado com sucesso!");
//		frame.setVisible(false);
//		//jb.setVisible(true);
//		JOptionPane.showMessageDialog(frame, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "CRM vs. MIS - Operações ativas", 0, null);
	    
//		final JDialog dialog = pane.createDialog(frame, "CRM vs. MIS");
////		//JOptionPane.showMessageDialog(null, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "", JOptionPane.OK_OPTION);
////		dialog.setVisible(true);
////		frame.setVisible(false);
//		
////		final JDialog dialog = pane.createDialog("Error");
//		
//		 dialog.addWindowListener(null);
//		 dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
//		 
//		 ScheduledExecutorService sch = Executors.newSingleThreadScheduledExecutor();     
//		 sch.schedule(new Runnable() {
//		     public void run() {
//		         dialog.setVisible(false);
//		         dialog.dispose();
//		     }
//		 }, 5, TimeUnit.SECONDS);
//		 
//		 dialog.setVisible(true);
		 
//		int i = 0;
//		while(i != 10) {
//			if(i == 0 ) {
//				
//			}
//			System.out.println(i);
//			i++;
//			try {
//				Thread.sleep(300);
//			} catch (InterruptedException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//		}
		
//		JOptionPane.showMessageDialog(frame, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "", JOptionPane.OK_OPTION);
		
		
//		Thread thread2 = new Thread(){
//			public void run() {
//			ImageIcon icon;
//			//URL img = this.getClass().getResource("loading.gif");
//			URL url = null;
//			try {
//				url = Paths.get("target", "loading.gif").toUri().toURL();
//			} catch (MalformedURLException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//			System.out.println(url);
//			//System.out.println(this.getClass().getResource(url));
//			icon = new ImageIcon(url);
//			JFrame frame = new JFrame("");
//			frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//			JOptionPane pane = new JOptionPane("Carregando depara das operações ativas...", JOptionPane.INFORMATION_MESSAGE, JOptionPane.DEFAULT_OPTION, null, new Object[]{}, null);
//			pane.setIcon(icon);
//			pane.setVisible(true);
////			JProgressBar jProgressBar = new JProgressBar(1, 100);
////			jProgressBar.setValue(15);
////			pane.add(jProgressBar,1);
//			JDialog dialog = pane.createDialog(frame, "CRM vs. MIS");
////			//JOptionPane.showMessageDialog(null, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "", JOptionPane.OK_OPTION);
//			dialog.setVisible(true);	
//		
//			}	
//			
//		};
//		
//		Thread thread1 = new Thread(){
//			public void run() {
//				JOptionPane.showMessageDialog(null, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "", JOptionPane.OK_OPTION);
//			}
//		};
	
//		thread2.start();
//		int i = 0;
//		while(i != 10) {
//			System.out.println(i);
//			i++;
//			try {
//				thread2.sleep(200);
//			} catch (InterruptedException e) {
//				// TODO Auto-generated catch block
//				e.printStackTrace();
//			}
//		}
////		thread2.interrupt();
////		thread2.suspend();
////		thread2.
////		while(thread2.isAlive()) {
////			System.out.println(i);
////			i++;
////			try {
////				thread2.sleep(500);
////			} catch (InterruptedException e) {
////				// TODO Auto-generated catch block
////				e.printStackTrace();
////			}
////		}
//		
//		thread1.start();
//		if(thread1.isAlive()) {
//			thread2.stop();
//		}
//jop.setVisible(false);
		
		
		
		
//		ImageIcon icon;
//		//URL img = this.getClass().getResource("loading.gif");
//		URL url = null;
//		try {
//			url = Paths.get("target", "loading.gif").toUri().toURL();
//		} catch (MalformedURLException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		System.out.println(url);
//		//System.out.println(this.getClass().getResource(url));
//		icon = new ImageIcon(url);
//		JFrame frame = new JFrame("");
////		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//		JOptionPane pane = new JOptionPane("Carregando depara das operações ativas...", JOptionPane.INFORMATION_MESSAGE, JOptionPane.DEFAULT_OPTION, null, new Object[]{}, null);
//		pane.setIcon(icon);
//		pane.setVisible(true);
////		JProgressBar jProgressBar = new JProgressBar(1, 100);
////		jProgressBar.setValue(15);
////		pane.add(jProgressBar,1);
//		JDialog dialog = pane.createDialog(frame, "CRM vs. MIS");
////		//JOptionPane.showMessageDialog(null, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "", JOptionPane.OK_OPTION);
//		dialog.setVisible(true);
////	    dialog.dispose();
////	    dialog = pane.createDialog(frame, "Information message");
//	    JOptionPane.showMessageDialog(null, "Depara CRM vs. MIS para operações ativas finalizado com sucesso!", "", JOptionPane.OK_OPTION);
		
//		listf("C:\\");
//		getFileList("C:\\");
		
//        File[] fileList = getFileList("C:\\");
//
//        //for(File file : fileList) {
//            //System.out.println(file.getName());
//            if(fileList != null)
//	        for (File file : fileList) {
//	        	
//	            if (file.isFile()) {
//	            	 System.out.println(file.getName());
//	            } else if (file.isDirectory()) {
//	            	System.out.println("aqui");
//	                fileList = getFileList(file.getAbsolutePath());      
//	            }
//	            
//	       }

       //}
		
//        File[] fileList = getFileList("C:\\");
//
//        for(File file : fileList) {
//            System.out.println(file.getName());
//        }
		
//		getPath("Directory.jar", "C:\\");
		
		
	}


	public File[] getFileList(String dirPath) {
		
        File dir = new File(dirPath);   
        //final File[] fileList2 = getFileList(dirPath);
        
        final File[] fileList = dir.listFiles(
        	new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.startsWith("Directory") && name.endsWith(".jar");
        }});
		return fileList; 
		
//		File dir = new File(dirPath);
//		final File[] fList = dir.listFiles();
//		final List<File> files = new ArrayList<File>();
//		File[] matches = dir.listFiles(new FilenameFilter()
//		{
//		  public boolean accept(File dir, String name)
//		  {
//      	    if(fList != null)
//  	        for (File file : fList) {      
//  	            if (file.isFile()) {
//  	            	//files.add(file);
//  	            	return name.startsWith("Directory") && name.endsWith(".jar");
//  	            } else if (file.isDirectory()) {
//  	                listf(file.getAbsolutePath());
//  	            }
//  	       }
//		     return name.startsWith("Directory") && name.endsWith(".jar");
//		  }
//		});
//		return matches;
//        
        //return name.startsWith("Directory") && name.endsWith(".jar");
        //return fileList;
	}
	
	public void listf(String directoryName) {
//	    File directory = new File(directoryName);
//	    List<File> files = new ArrayList<File>();
//	    // Get all files from a directory.
//	    File[] fList = directory.listFiles();
//	    if(fList != null)
//	        for (File file : fList) {      
//	            if (file.isFile()) {
//	                files.add(file);
//	            } else if (file.isDirectory()) {
//	                listf(file.getAbsolutePath());
//	            }
//	        }
		
		File f = new File(directoryName);
		
		File[] matchingFiles = f.listFiles(new FilenameFilter() {
		    public boolean accept(File dir, String name) {
		        return name.startsWith("Directory") && name.endsWith("jar");
		    }
		});
		//System.out.println(matchingFiles);
		//return matchingFiles;
	}
	
	public static String getPath(String drl, String whereIAm) {
	    File dir = new File(whereIAm); //StaticMethods.currentPath() + "\\src\\main\\resources\\" + 
	    for(File e : dir.listFiles()) {
	        if(e.isFile() && e.getName().equals(drl)) {return e.getPath();}
	        if(e.isDirectory()) {
	            String idiot = getPath(drl, e.getPath());
	            if(idiot != null) {return idiot;}
	        }
	    }
	    return null;
	}

}
