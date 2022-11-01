package br.com.misvscrm.utils;

import java.awt.Toolkit;

import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;

public class InterfaceView extends JFrame{
	
	JProgressBar jb;
	
	public InterfaceView(String titulo) {
		
		super(titulo);
		jb = new JProgressBar(0, 100);
		
		jb.setBounds(40,40,160,30);   
		jb.setValue(0);    
		jb.setStringPainted(true); 
		add(jb);    
		setSize(getTitle().length() * 12, 150);    
		setLayout(null);
		setLocationRelativeTo(null);
		setVisible(true);
		
	}
	
	public void setProgressBarValue(int value) {
		jb.setValue(value);
	}
	
	public void closeFrame() {
		setVisible(false);
	}
	
	public void showMessage(String message) {
		Toolkit.getDefaultToolkit().beep();
		//JOptionPane.showMessageDialog(null, message, "", JOptionPane.OK_OPTION);
		JDialog dialog = new JDialog();
		dialog.setAlwaysOnTop(true);
		JOptionPane.showMessageDialog(dialog, message);
		//System.exit(0);
	}
	
	public JFrame newBlankFrame(String title) {
		JFrame frame = new JFrame(title);
		frame.setSize(getTitle().length() * 12, 150);    
		frame.setLayout(null);
		frame.setLocationRelativeTo(null);
		frame.setVisible(true);
		return frame;
	}
	
	public void systemExit() {
		System.exit(0);
	}
}
