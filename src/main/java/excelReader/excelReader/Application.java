package excelReader.excelReader;

import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Frame;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Iterator;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import java.awt.Color;

import javax.swing.JTextField;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JLabel;
import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.swing.ImageIcon;

public class Application extends Frame {

	private JFrame frame;

	/**
	 * * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Application window = new Application();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public Application() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 600, 400);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		addIcrisatLogo();

		JLabel lblFile = new JLabel("File Name");
		lblFile.setFont(new Font("Arial", Font.BOLD | Font.ITALIC, 20));
		lblFile.setBounds(10, 141, 121, 20);
		frame.getContentPane().add(lblFile);

		final JTextField textField = new JTextField();
		textField.setBounds(141, 144, 334, 20);
		frame.getContentPane().add(textField);
		textField.setColumns(10);

		JButton btnSubmit = new JButton("Submit");
		btnSubmit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				convert(textField.getText());
			}
		});
		btnSubmit.setBounds(256, 191, 89, 23);
		frame.getContentPane().add(btnSubmit);

		JLabel lblFileConverter = new JLabel("File Converter");
		lblFileConverter.setFont(new Font("Tahoma", Font.BOLD, 25));
		lblFileConverter.setForeground(Color.BLUE);
		lblFileConverter.setBounds(195, 32, 210, 44);
		frame.getContentPane().add(lblFileConverter);
		
		JLabel lblhornysaze = new JLabel("@hornySaze");
		lblhornysaze.setBounds(10, 336, 101, 14);
		frame.getContentPane().add(lblhornysaze);
		
		JButton btnBrowse = new JButton("Browse");
		btnBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser fileChooser = new JFileChooser();
				buttonActionPerformed(arg0, textField,fileChooser);      	
			}
		});
		btnBrowse.setBounds(485, 143, 89, 23);
		frame.getContentPane().add(btnBrowse);

	}

	private void buttonActionPerformed(ActionEvent evt, JTextField fileName, JFileChooser fileChooser) {
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            fileName.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
}

	
	public static void convert(String fileName) {
		String outputFile = fileName.substring(0, fileName.lastIndexOf(".")) + "_processed"
				+ fileName.substring(fileName.lastIndexOf("."));
		XSSFWorkbook outputWorkbook = new XSSFWorkbook();
		XSSFSheet sheet = outputWorkbook.createSheet("Datatypes in Java");

		try {
			FileInputStream excelFile = new FileInputStream(new File(fileName));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			int rowNum = 0;
			while (iterator.hasNext()) {
				Row inputRow = iterator.next();
				Iterator<Cell> inputCellIterator = inputRow.iterator();

				int colNum = 0;
				Cell inputCell;
				while (inputCellIterator.hasNext()) {
					Row outputRow = sheet.createRow(rowNum++);
					inputCell = inputCellIterator.next();
					addCell(colNum, inputCell, outputRow);
					colNum++;
					inputCell = inputCellIterator.next();
					addCell(colNum, inputCell, outputRow);
					colNum = 0;
				}
			}

			FileOutputStream outputStream = new FileOutputStream(outputFile);
			outputWorkbook.write(outputStream);
			outputWorkbook.close();
			workbook.close();
		} catch (FileNotFoundException e1) {
			e1.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("deprecation")
	private static void addCell(int colNum, Cell inputCell, Row outputRow) {
		Cell outputCell = outputRow.createCell(colNum++);
		if (inputCell.getCellTypeEnum() == CellType.STRING) {
			outputCell.setCellValue(inputCell.getStringCellValue());
		} else if (inputCell.getCellTypeEnum() == CellType.NUMERIC) {
			outputCell.setCellValue(inputCell.getNumericCellValue());
		}
	}

	public void addIcrisatLogo() {
		JLabel icrisatLogoLabel = new JLabel("");
		
		URL url = Application.class.getResource("icrisat.png");
		if(url != null){
			ImageIcon icon = new ImageIcon(new ImageIcon(url).getImage().getScaledInstance(150, 50, Image.SCALE_DEFAULT));
			icrisatLogoLabel.setIcon(icon);
			icrisatLogoLabel.setBounds(407, 286, 167, 50);
			frame.getContentPane().add(icrisatLogoLabel);
		}
	}
}
