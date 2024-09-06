package com.github.st0rm1O1;
import com.github.st0rm1O1.frame.ApplicationFrame;

import javax.swing.JOptionPane;

public class ExcelReader {

	public static void main(String[] args) {
		
		try { new ApplicationFrame(); }
		
		catch (Throwable t) {
			JOptionPane.showMessageDialog(null, t.getClass().getSimpleName() + " : " + t.getMessage());
			throw t;
		} // catch()
		
	}
}
