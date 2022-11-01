package br.com.misvscrm;

import java.io.File;
import java.io.FilenameFilter;

public class MyFilenameFilter implements FilenameFilter{

    String initials;
    
    // constructor to initialize object
    public MyFilenameFilter(String initials)
    {
        this.initials = initials;
    }
    
    // overriding the accept method of FilenameFilter
    // interface
    public boolean accept(File dir, String name)
    {
        return name.startsWith(initials);
    }
       
//	File directory = new File("src/planilhas/operações em aberto");
//	MyFilenameFilter filter = new MyFilenameFilter("Oportunidades Abertas 14-09-2022 18-05-02.xlsx");
//    String[] flist = directory.list(filter);	
//	System.out.println(flist[0]);
//	File file = new File(flist[0]);
    
}
