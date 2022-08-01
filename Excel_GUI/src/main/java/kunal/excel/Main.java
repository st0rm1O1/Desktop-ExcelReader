package kunal.excel;
import javax.swing.JOptionPane;

public class Main {

	public static void main(String[] args) {
		
		try { new GUI(); } 
		
		catch (Throwable t) {
			JOptionPane.showMessageDialog(null, t.getClass().getSimpleName() + " : " + t.getMessage());
			throw t;
		} // catch()
		
	}
}
