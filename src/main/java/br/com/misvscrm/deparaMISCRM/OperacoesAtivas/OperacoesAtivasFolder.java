package br.com.misvscrm.deparaMISCRM.OperacoesAtivas;

import java.io.File;
import java.net.URISyntaxException;

import br.com.misvscrm.utils.FileUtils;
import br.com.misvscrm.utils.InterfaceView;

public class OperacoesAtivasFolder {

	public static void main(String[] args) {
		
		try {
			System.out.println(new File(OperacoesAtivasFolder.class.getProtectionDomain().getCodeSource().getLocation()
				    .toURI()).getParent());
			System.out.println("aqui" + System.getProperty("user.dir"));
		} catch (URISyntaxException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		FileUtils fileUtils = new FileUtils();

		InterfaceView interfaceView = new InterfaceView("Criando pastas para Operações ativas");
		interfaceView.setProgressBarValue(50);
		fileUtils.createFolder("Planilhas/Operações Ativas2");
		interfaceView.setProgressBarValue(100);
		interfaceView.closeFrame();
		interfaceView.showMessage("Pastas criadas com sucesso!");
		interfaceView.showMessage("Dentro do diretório da 'ÁREA DE TRABALHO', foi criado o caminho 'Planilhas/Operações Ativas'.\n"
								  + "Insira as planilhas geradas no CRM e no MIS e execute o arquivo DeparaOperaçõesAtivas.jar.");
		interfaceView.systemExit();
		
	}
		
}
