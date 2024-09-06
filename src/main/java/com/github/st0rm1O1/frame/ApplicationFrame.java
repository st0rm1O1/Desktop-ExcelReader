package com.github.st0rm1O1.frame;

import java.util.Arrays;
import java.awt.Insets;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPopupMenu;
import javax.swing.BorderFactory;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JTextArea;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import com.github.st0rm1O1.common.FileIO;
import com.github.st0rm1O1.resource.Resource;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;





public class ApplicationFrame extends JFrame {

	private static final long serialVersionUID = 1L;
	
	private JLabel txt, panel, op_txt;
	private JButton file_btn, crud_btn, op_btn, close_btn;
	private JTextField field;
	private JTextArea path_txt;
	private JTable display_table;
	private DefaultTableModel tableModel;
	private JScrollPane scroll_dis, scroll_path;
	private JPopupMenu pop_crud, pop_insert;
	private FileIO util;
	
	private File file;
	private FileInputStream fileInput;
	private FileOutputStream fileOutput;
	private XSSFWorkbook workBook;
	private XSSFSheet sheet;
	
	private int row_index, col_index;
	private Object[] col_obj, row_obj;

	
	
	public ApplicationFrame() {
		initialize();
	}
	
	private void readExcel() {
		
		try {
			
			if (file.length() != 0) {
				
				fileInput = new FileInputStream(file);
				workBook = new XSSFWorkbook(fileInput);
				sheet = workBook.getSheetAt(0);
				
				col_obj = new Object[sheet.getRow(0).getLastCellNum()];
				row_obj = new Object[sheet.getLastRowNum()];
				
				for(int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
					col_obj[i] = sheet.getRow(0).getCell(i).getStringCellValue();
				}
				
				tableModel.setColumnIdentifiers(col_obj);
				display_table.setModel(tableModel);
				
				
				for(int i = 1; i <= sheet.getLastRowNum(); i++) {
					
					for(int j = 0; j < sheet.getRow(i).getLastCellNum(); j++) {
						
						Cell cell = sheet.getRow(i).getCell(j);
						
						if (cell != null) {
						
							switch (cell.getCellType()) {
								
									case NUMERIC:
//										System.out.println(sheet.getRow(i).getCell(j).getNumericCellValue());
										double d = cell.getNumericCellValue();
										
										if (d % 1 == 0)
											row_obj[j] = Math.round(d);
										
										else
											row_obj[j] = d;
										
										break;
										
									case STRING:
//										System.out.println(sheet.getRow(i).getCell(j).getStringCellValue());
										row_obj[j] = cell.getStringCellValue();
										break;
										
									case BLANK:
										System.out.printf("BLANK AT (%d ROW) - (%d COL)\n", i, j+1);
										row_obj[j] = null;
										break;
										
									case BOOLEAN:
//										System.out.println(sheet.getRow(i).getCell(j).getBooleanCellValue());
										row_obj[j] = cell.getBooleanCellValue();
										break;
										
									case ERROR:
										System.out.printf("ERROR AT (%d ROW) - (%d COL)\n", i, j+1);
										row_obj[j] = "ERROR";
										break;
										
									default:
										System.out.printf("DEFAULT CELL (%d ROW) - (%d COL)\n", i, j+1);
										row_obj[j] = "DEFAULT";
										break;
										
							} // switch
						
						} // if null
						
					} // for j
					
					tableModel.addRow(row_obj);
					Arrays.fill(row_obj, null);
					
				} // for i
				
			} // file length
			
		} catch (Exception e) { e.printStackTrace(); }
		
	}
	
	private void expandUtil(String title, boolean val) {
		
		if (workBook == null) 
			JOptionPane.showMessageDialog(this, "NO FILE SELECTED!", "E404 (Not Found)", JOptionPane.ERROR_MESSAGE);

		else {
		
			op_txt.setText(title);
			op_btn.setText(title);
			field.setVisible(val);
			field.setText(null);
			this.setSize(1100, 600);
						
		}
	}
	
	private void initialize() {
		
		util = new FileIO();
		
		UIManager.put("OptionPane.messageFont", Resource.getInterRegular(16));
		UIManager.put("OptionPane.buttonFont", Resource.getInterMedium(15));
		UIManager.put("OptionPane.buttonAreaBorder", null);
		
		setBounds(100, 100, 700, 600);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setTitle("Excel Viewer - st0rm1O1");
		setResizable(false);
		setIconImage(Resource.loadImage(Resource.ICON_PATH_ICON));
		getContentPane().setBackground(new Color(245, 245, 245));
		getContentPane().setLayout(null); 
		
		txt = new JLabel("Location -");
		txt.setVerticalAlignment(SwingConstants.TOP);
		txt.setFont(Resource.getInterSemibold(18));
		txt.setBounds(20, 20, 180, 30);


		path_txt = new JTextArea("No File Selected.");
		scroll_path = new JScrollPane(path_txt, JScrollPane.VERTICAL_SCROLLBAR_NEVER, JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		scroll_path.setBorder(null);
		scroll_path.setBounds(20, 50, 480, 50);
		
		path_txt.setBackground(new Color(245, 245, 245));
		path_txt.setForeground(Color.BLACK);
		path_txt.setEditable(false);
		path_txt.setFont(Resource.getInterRegular(16));
		
		
		display_table = new JTable();
		tableModel = new DefaultTableModel() {

			private static final long serialVersionUID = 1L;

			@Override
		    public boolean isCellEditable(int row, int column) {
		       return false;
		    }
		};
		scroll_dis = new JScrollPane(display_table, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		scroll_dis.setBorder(null);
		scroll_dis.setBounds(20, 135, 650, 410);
		
		display_table.setFocusable(false);
		display_table.setSelectionBackground(new Color(255, 240, 245));
		display_table.setRowMargin(2);
		display_table.setRowHeight(25);
		display_table.getTableHeader().setPreferredSize(new Dimension(scroll_dis.getWidth(), 30));
		display_table.getTableHeader().setFont(Resource.getInterMedium(18));
		display_table.setFont(Resource.getInterRegular(16));
		display_table.setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
		((JLabel) display_table.getDefaultRenderer(Object.class)).setHorizontalAlignment(SwingConstants.CENTER);
		
		display_table.addMouseListener(new MouseAdapter() {
			@Override
			public void mousePressed(MouseEvent e) {
				if ("UPDATE".equals(op_txt.getText())) {
					row_index = display_table.getSelectedRow();
					col_index = display_table.getSelectedColumn();
					Object objValue = display_table.getModel().getValueAt(row_index, col_index);
					if (objValue != null) {
						field.setText(objValue.toString());
					} else {
						field.setText(null);
					}
				}
			}
		});

		
		
		pop_crud = new JPopupMenu();
		pop_insert = new JPopupMenu();
		
		pop_crud.setBorder(BorderFactory.createCompoundBorder(pop_crud.getBorder(),BorderFactory.createEmptyBorder(8, 0, 8, 0)));
		pop_insert.setBorder(BorderFactory.createCompoundBorder(pop_crud.getBorder(),BorderFactory.createEmptyBorder(4, 0, 4, 0)));

		JMenuItem c_op = new JMenuItem("Create Excel", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_CREATE)));
	    JMenuItem i_op = new JMenuItem("New Record", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_INSERT)));
	    JMenuItem u_op = new JMenuItem("Update Record", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_UPDATE)));
	    JMenuItem d_op = new JMenuItem("Delete Record", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_DELETE)));
	    JMenuItem up_in = new JMenuItem("Insert - Top", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_TOP)));
	    JMenuItem down_in = new JMenuItem("Insert - Bottom", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_BOTTOM)));
	    JMenuItem left_in = new JMenuItem("Insert - Left", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_LEFT)));
	    JMenuItem right_in = new JMenuItem("Insert - Right", new ImageIcon(Resource.loadImage(Resource.ICON_PATH_RIGHT)));
	    
	    c_op.setIconTextGap(10);
	    i_op.setIconTextGap(10);
	    u_op.setIconTextGap(10);
	    d_op.setIconTextGap(10);
	    up_in.setIconTextGap(10);
	    down_in.setIconTextGap(10);
	    left_in.setIconTextGap(10);
	    right_in.setIconTextGap(10);
	    

	    c_op.setFont(Resource.getInterRegular(16));
	    i_op.setFont(Resource.getInterRegular(16));
	    u_op.setFont(Resource.getInterRegular(16));
	    d_op.setFont(Resource.getInterRegular(16));
	    up_in.setFont(Resource.getInterRegular(16));
	    down_in.setFont(Resource.getInterRegular(16));
	    left_in.setFont(Resource.getInterRegular(16));
	    right_in.setFont(Resource.getInterRegular(16));
	    
	    pop_crud.add(c_op);
	    pop_crud.add(i_op);
	    pop_crud.add(u_op);
	    pop_crud.add(d_op);
	    
	    pop_insert.add(up_in);
	    pop_insert.add(down_in);
	    pop_insert.add(left_in);
	    pop_insert.add(right_in);
	   
	    
	    c_op.addActionListener(e -> util.createExcel(new JFileChooser().getFileSystemView().getDefaultDirectory().toString(), JOptionPane.showInputDialog(ApplicationFrame.this, "File-Name", "Create Excel Sheet", JOptionPane.INFORMATION_MESSAGE)));
	    i_op.addActionListener(e -> expandUtil("INSERT", false));
	    u_op.addActionListener(e -> expandUtil("UPDATE", true));
	    d_op.addActionListener(e -> expandUtil("DELETE", false));
	    
	    up_in.addActionListener(e -> {

            row_index = display_table.getSelectedRow();

            if (row_index == -1)
                JOptionPane.showMessageDialog(ApplicationFrame.this, "NO ROW SELECTED!", "500 (Internal System Error)", JOptionPane.ERROR_MESSAGE);

            else {

                util.insertRecordExcel(row_index+1, 1, file);
                tableModel.setColumnCount(0);
                tableModel.setRowCount(0);
                tableModel.fireTableDataChanged();
                readExcel();
            }

        });
	    
	    down_in.addActionListener(e -> {

            row_index = display_table.getSelectedRow();

            if (row_index == -1)
                JOptionPane.showMessageDialog(ApplicationFrame.this, "NO ROW SELECTED!", "500 (Internal System Error)", JOptionPane.ERROR_MESSAGE);

            else {

                util.insertRecordExcel(row_index+2, 1, file);
                tableModel.setColumnCount(0);
                tableModel.setRowCount(0);
                tableModel.fireTableDataChanged();
                readExcel();
            }

        });
	    

		
		file_btn = new JButton("OPEN");
		file_btn.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
		file_btn.setBounds(520, 20, 140, 40);
		file_btn.setAlignmentX(Component.CENTER_ALIGNMENT);
		file_btn.setPreferredSize(new Dimension(200, 40));
		file_btn.setForeground(Color.WHITE);
		file_btn.setFocusPainted(false);
		file_btn.setBackground(new Color(255, 51, 102));
		file_btn.setBorder(null);
		file_btn.setFont(Resource.getInterSemibold(18));
		file_btn.addActionListener(a -> {

            try {

                if (fileInput != null)
                    fileInput.close();

                if (fileOutput != null)
                    fileOutput.close();

                if (workBook != null)
                    workBook.close();

            } catch (Exception e) { e.printStackTrace(); }


            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setCurrentDirectory(new File("."));
            fileChooser.setFileFilter(new FileNameExtensionFilter("Excel Sheet (.xlsx)", "xlsx", "excel"));
            int response = fileChooser.showOpenDialog(null);

            if (response == JFileChooser.APPROVE_OPTION) {

                tableModel.setColumnCount(0);
                tableModel.setRowCount(0);
                tableModel.fireTableDataChanged();
                path_txt.setText("No File Selected.");
                op_txt.setText("CRUD");
                op_btn.setText("OPERATION");
                ApplicationFrame.this.setSize(700, 600);

                file = new File(fileChooser.getSelectedFile().getAbsoluteFile().toString());
                path_txt.setText(file.getPath());

                readExcel();

            } // if - success


        });
		
		crud_btn = new JButton("CRUD");
		crud_btn.setPreferredSize(new Dimension(200, 40));
		crud_btn.setForeground(Color.WHITE);
		crud_btn.setFont(Resource.getInterRegular(18));
		crud_btn.setFocusPainted(false);
		crud_btn.setBorder(null);
		crud_btn.setBackground(new Color(0, 102, 255));
		crud_btn.setAlignmentX(0.5f);
		crud_btn.setBounds(520, 70, 140, 40);
		crud_btn.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				pop_crud.show(crud_btn, e.getX(), e.getY());
			}
		});
		
		op_txt = new JLabel("CRUD");
		op_txt.setHorizontalAlignment(SwingConstants.CENTER);
		op_txt.setForeground(Color.BLACK);
		op_txt.setFont(Resource.getInterSemibold(40));
		op_txt.setBounds(720, 120, 320, 60);
		
		field = new JTextField();
		field.setMargin(new Insets(2, 10, 2, 10));
		field.setForeground(Color.BLACK);
		field.setFont(Resource.getInterRegular(20));
		field.setBounds(750, 220, 250, 40);
		field.setColumns(10);
		
		panel = new JLabel();
		panel.setBackground(Color.WHITE);
		panel.setBorder(UIManager.getBorder("TextField.border"));
		panel.setBounds(720, 50, 320, 450);
		
		op_btn = new JButton("OPERATION");
		op_btn.setPreferredSize(new Dimension(200, 40));
		op_btn.setForeground(Color.WHITE);
		op_btn.setFont(Resource.getInterSemibold(18));
		op_btn.setFocusPainted(false);
		op_btn.setBorder(null);
		op_btn.setBackground(new Color(51, 102, 153));
		op_btn.setAlignmentX(0.5f);
		op_btn.setBounds(750, 340, 250, 40);
		op_btn.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				if ("INSERT".equals(op_txt.getText())) {
					pop_insert.show(op_btn, e.getX(), e.getY());
				}
			}
		});
		
		op_btn.addActionListener(e -> {
                if ("DELETE".equals(op_txt.getText())) {

                    row_index = display_table.getSelectedRow();

                    if (row_index == -1)
                        JOptionPane.showMessageDialog(ApplicationFrame.this, "NO ROW SELECTED!", "500 (Internal System Error)", JOptionPane.ERROR_MESSAGE);

                    else {

                        util.deleteRecordExcel(row_index+1, 1, file);
                        tableModel.setColumnCount(0);
                        tableModel.setRowCount(0);
                        tableModel.fireTableDataChanged();
                        readExcel();
                    }

                } // DELETE

                if ("UPDATE".equals(op_txt.getText())) {

                    row_index = display_table.getSelectedRow();
                    col_index = display_table.getSelectedColumn();

                    if (row_index == -1)
                        JOptionPane.showMessageDialog(ApplicationFrame.this, "NO ROW SELECTED!", "500 (Internal System Error)", JOptionPane.ERROR_MESSAGE);

                    else {

                        util.updateRecordExcel(row_index+1, col_index, file, field.getText());
                        tableModel.setColumnCount(0);
                        tableModel.setRowCount(0);
                        tableModel.fireTableDataChanged();
                        readExcel();
                    }


                } // UPDATE
        });

		
		close_btn = new JButton("CLOSE");
		close_btn.setPreferredSize(new Dimension(200, 40));
		close_btn.setForeground(Color.WHITE);
		close_btn.setFont(Resource.getInterSemibold(18));
		close_btn.setFocusPainted(false);
		close_btn.setBorder(null);
		close_btn.setBackground(new Color(102, 102, 204));
		close_btn.setAlignmentX(0.5f);
		close_btn.setBounds(750, 400, 250, 40);
		close_btn.addActionListener(a -> {
            op_txt.setText("CRUD");
            op_btn.setText("OPERATION");
            field.setVisible(false);
            field.setText(null);
            ApplicationFrame.this.setSize(700, 600);
        });

		
		getContentPane().add(txt);
	    	getContentPane().add(scroll_path);
	    	getContentPane().add(scroll_dis);
		getContentPane().add(file_btn);
		getContentPane().add(crud_btn);
		getContentPane().add(pop_crud);
		getContentPane().add(op_txt);
		getContentPane().add(field);
		getContentPane().add(op_btn);
		getContentPane().add(close_btn);
		getContentPane().add(panel);
		setVisible(true);
		
	}
}
