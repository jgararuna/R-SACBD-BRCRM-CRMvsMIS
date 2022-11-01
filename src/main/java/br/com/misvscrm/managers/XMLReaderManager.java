package br.com.misvscrm.managers;

import java.io.File;
import java.io.IOException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

public class XMLReaderManager {
	
	DocumentBuilder db = null;
	Document doc = null;
	public XMLReaderManager() {
		
		File file = new File("config//DirectoryConfiguration.xml"); 
		//misvscrm/config/DirectoryConfiguration.xml
		//File file = new File("/misvscrm/config/DirectoryConfiguration.xml"); 
		//an instance of factory that gives a document builder  
		DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();  
		//an instance of builder to parse the specified xml file  
		//DocumentBuilder db = null;
		try {
			db = dbf.newDocumentBuilder();
		} catch (ParserConfigurationException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}  
		//Document doc = null;
		try {
			doc = db.parse(file);
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  
		doc.getDocumentElement().normalize();
		//System.out.println("Root element: " + doc.getDocumentElement().getNodeName());
	}
	
	public static void main(String[] args) {
		XMLReaderManager xmlReaderManager = new XMLReaderManager();
		xmlReaderManager.getNodeList("OperaçõesAtivas");
		xmlReaderManager.getUrlBaseOperacoesAtivas();
	}
	
	public Element getNodeList(String nodeName) {
		
		NodeList nodeList = doc.getElementsByTagName(nodeName);
		//System.out.println(nodeList.getLength());
		Element eElement = null;
		for (int itr = 0; itr < nodeList.getLength(); itr++)   
		{  
			Node node = nodeList.item(itr);  
			//System.out.println("\nNode Name :" + node.getNodeName());  
			if (node.getNodeType() == Node.ELEMENT_NODE)   
			{  
				eElement = (Element) node; 
				
			}
		}
		return eElement;	
		
	}
	
	/*
	 * Dados operações ativas
	 */
	public String getNodeElementText(String tipoDepara, String node) {
		Element element = getNodeList(tipoDepara);
		String nodeTextContent = element.getElementsByTagName(node).item(0).getTextContent();
		return nodeTextContent;
	}
	public String getUrlBaseOperacoesAtivas() {	
		Element element = getNodeList("OperaçõesAtivas");
		String urlBaseOperacoesAtivas = element.getElementsByTagName("urlBase").item(0).getTextContent();
		return urlBaseOperacoesAtivas;
	}
	
	public String getCRMOperacoesAtivas() {		
		Element element = getNodeList("OperaçõesAtivas");
		String CRMOperacoesAtivas = element.getElementsByTagName("fileNameCRM").item(0).getTextContent();
		return CRMOperacoesAtivas;
	}
	
	public String getMISOperacoesAtivas() {
		Element element = getNodeList("OperaçõesAtivas");
		String MISOperacoesAtivas = element.getElementsByTagName("fileNameMIS").item(0).getTextContent();
		return MISOperacoesAtivas;
	}
	
	public String getRenamedCRMOperacoesAtivas() {		
		Element element = getNodeList("OperaçõesAtivas");
		String RenamedCRMOperacoesAtivas = element.getElementsByTagName("fileRenameCRM").item(0).getTextContent();
		return RenamedCRMOperacoesAtivas;
	}
	
	public String getRenamedMISOperacoesAtivas() {
		Element element = getNodeList("OperaçõesAtivas");
		String RenamedMISOperacoesAtivas = element.getElementsByTagName("fileRenameMIS").item(0).getTextContent();
		return RenamedMISOperacoesAtivas;
	}
	
}
