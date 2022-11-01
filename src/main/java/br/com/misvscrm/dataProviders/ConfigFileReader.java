package br.com.misvscrm.dataProviders;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.Properties;

public class ConfigFileReader {
	
	private Properties prop;
	private final String local = "config//FolderConfiguration.properties";
	String urlBase;
	public ConfigFileReader() {
		BufferedReader reader;

		try {
			reader = new BufferedReader(new FileReader(local));
			prop = new Properties();

			try {
				prop.load(reader);
				reader.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new RuntimeException("Configuracao.properties não encontrado em " + local);
		}
		
		urlBase = prop.getProperty("urlBase");
		
	}
	
	public String getUrlBase() {
		String url = urlBase;
		return url;
	}
	
	public String getUrlOperacoesEmAberto() {
		//String url = urlBase + prop.getProperty("urlOperacoesEmAberto");
		String url = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + "Planilhas" + File.separator + "Operações Ativas";
		return url;
	}
	
	public String getfileNameOperacoesEmAbertoCRM() {
		String fileName = prop.getProperty("fileNameOperacoesEmAbertoCRM");
		return fileName;
	}
	
	public String getfileNameOperacoesEmAbertoMIS() {
		String fileName = prop.getProperty("fileNameOperacoesEmAbertoMIS");
		return fileName;
	}

}
