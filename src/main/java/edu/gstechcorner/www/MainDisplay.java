package edu.gstechcorner.www;

import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;
import org.eclipse.wb.swt.SWTResourceManager;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.eclipse.swt.widgets.Label;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.List;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Point;

public class MainDisplay {

	protected Shell shell;
	private Text txtCustomerName;
	private Text txtUPC;
	private Text txtEagleID;
	private Text txtEmail;
	private Text txtPhoneNum;
	private Label lblDate;
	private Label lblProductUPC;
	private Label lblStatus;
	private Text textUPC2;
	private Text textUPC3;
	private Text textUPC4;
	private Label lblDescription;
	private Label lblUnitPrice;
	private Text textDescription1;
	private Text textDescription2;
	private Text textDescription3;
	private Text textDescription4;
	private Text textPrice1;
	private Text textPrice2;
	private Text textPrice3;
	private Text textPrice4;
	private Text txtTotal;
	private Text txtSubtotal;
	private Text txtTax;
	private Label lblSerialNumber;
	private Text textSerial1;
	private Text textSerial2;
	private Text textSerial3;
	private Text textSerial4;
	private Label labelQuantity;
	private Text textQuantity1;
	private Text textQuantity2;
	private Text textQuantity3;
	private Text textQuantity4;
	private Label labelPaymentType;
	private Text txtpayment1;
	private Text txtpayment2;
	private Text txtpayment3;
	private Text txtpayment4;
	
		/**
	 * Launch the application.
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			MainDisplay window = new MainDisplay();
			window.open();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Open the window.
	 */
	public void open() {
		Display display = Display.getDefault();
		createContents();
		shell.open();
		shell.layout();
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}

	/**
	 * Create contents of the window.
	 */
	protected void createContents() {

		shell = new Shell();
		shell.setMinimumSize(new Point(173, 52));
		shell.setBackground(SWTResourceManager.getColor(SWT.COLOR_LIST_BACKGROUND));
		shell.setSize(1200,660);
		//shell.setSize(734, 576);
		shell.setText("GSU TechCorner");
		
		//Title
		Label lblTitle = new Label(shell, SWT.NONE);
		lblTitle.setFont(SWTResourceManager.getFont("Segoe UI", 20, SWT.BOLD));
		lblTitle.setBounds(20, 1, 381, 40);
		lblTitle.setText("GSU TechCorner");
		
		//Version
		Label lblVersion = new Label(shell, SWT.NONE);
		lblVersion.setBounds(1069, 21, 85, 20);
		lblVersion.setText("Version 1.0");
		
		//Status
		lblStatus = new Label(shell, SWT.BORDER);
		lblStatus.setBounds(1003, 44, 151, 20);
		lblStatus.setText("Status: " + "waiting for input");
		lblStatus.setFont(SWTResourceManager.getFont("Times New Roman", 8, SWT.BOLD));
				
		//Date
		lblDate = new Label(shell, SWT.NONE);
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM/dd/yyy");  
		LocalDateTime now = LocalDateTime.now();
		String Date = dtf.format(now);
		lblDate.setText(Date);
		lblDate.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblDate.setBounds(1052, 70, 105, 21);
	
		// Customer Name
		Label lblCustomer = new Label(shell, SWT.NONE);
		lblCustomer.setText("Customer Name");
		lblCustomer.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblCustomer.setBounds(20, 63, 151, 21);
		txtCustomerName = new Text(shell, SWT.BORDER);
		txtCustomerName.setBounds(20, 90, 151, 26);
		
		// Eagle ID
		Label lblEagleID = new Label(shell, SWT.NONE);
		lblEagleID.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblEagleID.setBounds(232, 63, 105, 26);
		lblEagleID.setText("Eagle ID");
		txtEagleID = new Text(shell, SWT.BORDER);
		txtEagleID.setBounds(232, 90, 133, 26);
		// Email
		Label lblEmail = new Label(shell, SWT.NONE);
		lblEmail.setText("Email");
		lblEmail.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblEmail.setBounds(20, 148, 105, 21);
		txtEmail = new Text(shell, SWT.BORDER);
		txtEmail.setBounds(20, 175, 151, 26);
		
		// Phone Number
		Label lblPhoneNum = new Label(shell, SWT.NONE);
		lblPhoneNum.setText("Phone #");
		lblPhoneNum.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblPhoneNum.setBounds(232, 148, 105, 21);
		txtPhoneNum = new Text(shell, SWT.BORDER);
		txtPhoneNum.setBounds(232, 175, 133, 25);
		
		// Products UPC
		lblProductUPC = new Label(shell, SWT.NONE);
		lblProductUPC.setText("Product upc");
		lblProductUPC.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblProductUPC.setBounds(20, 220, 172, 26);
		
		txtUPC = new Text(shell, SWT.BORDER);
		txtUPC.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		txtUPC.setBounds(20, 252, 182, 32);
		
		textUPC2 = new Text(shell, SWT.BORDER);
		textUPC2.setText("");
		textUPC2.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textUPC2.setBounds(20, 290, 182, 32);
		
		textUPC3 = new Text(shell, SWT.BORDER);
		textUPC3.setText("");
		textUPC3.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textUPC3.setBounds(20, 328, 182, 32);
		
		textUPC4 = new Text(shell, SWT.BORDER);
		textUPC4.setText("");
		textUPC4.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textUPC4.setBounds(20, 366, 182, 32);
		
		// Product Description
		lblDescription = new Label(shell, SWT.NONE);
		lblDescription.setText("Description");
		lblDescription.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblDescription.setBounds(236, 220, 286, 26);
		
		textDescription1 = new Text(shell, SWT.BORDER);
		textDescription1.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textDescription1.setBounds(232, 252, 290, 32);
		
		textDescription2 = new Text(shell, SWT.BORDER);
		textDescription2.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textDescription2.setBounds(232, 290, 290, 32);
		
		textDescription3 = new Text(shell, SWT.BORDER);
		textDescription3.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textDescription3.setBounds(232, 328, 290, 32);
		
		textDescription4 = new Text(shell, SWT.BORDER);
		textDescription4.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		textDescription4.setBounds(232, 366, 290, 32);
		
		//Unit Price
		lblUnitPrice = new Label(shell, SWT.NONE);
		lblUnitPrice.setText("Unit Price");
		lblUnitPrice.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblUnitPrice.setBounds(882, 220, 105, 26);
		
		textPrice1 = new Text(shell, SWT.BORDER);
		textPrice1.setText("0");
		textPrice1.setBounds(880, 252, 107, 32);
		
		textPrice2 = new Text(shell, SWT.BORDER);
		textPrice2.setText("0");
		textPrice2.setBounds(880, 290, 107, 32);
		
		textPrice3 = new Text(shell, SWT.BORDER);
		textPrice3.setText("0");
		textPrice3.setBounds(880, 328, 107, 32);
		
		textPrice4 = new Text(shell, SWT.BORDER);
		textPrice4.setText("0");
		textPrice4.setBounds(880, 366, 107, 32);
		
		// Serical Number
		lblSerialNumber = new Label(shell, SWT.NONE);
		lblSerialNumber.setText("Serial Number");
		lblSerialNumber.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		lblSerialNumber.setBounds(554, 220, 151, 26);
		
		textSerial1 = new Text(shell, SWT.BORDER);
		textSerial1.setText("N/A");
		textSerial1.setBounds(554, 252, 182, 32);
		
		textSerial2 = new Text(shell, SWT.BORDER);
		textSerial2.setText("N/A");
		textSerial2.setBounds(554, 290, 182, 32);
		
		textSerial3 = new Text(shell, SWT.BORDER);
		textSerial3.setText("N/A");
		textSerial3.setBounds(554, 328, 182, 32);
		
		textSerial4 = new Text(shell, SWT.BORDER);
		textSerial4.setText("N/A");
		textSerial4.setBounds(554, 366, 182, 32);
		
		//Quantity Count
		labelQuantity = new Label(shell, SWT.NONE);
		labelQuantity.setText("Quantity");
		labelQuantity.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		labelQuantity.setBounds(755, 220, 85, 26);
		
		textQuantity1 = new Text(shell, SWT.BORDER);
		textQuantity1.setText("0");
		textQuantity1.setBounds(755, 252, 85, 32);
		
		textQuantity2 = new Text(shell, SWT.BORDER);
		textQuantity2.setText("0");
		textQuantity2.setBounds(755, 290, 85, 32);
		
		textQuantity3 = new Text(shell, SWT.BORDER);
		textQuantity3.setText("0");
		textQuantity3.setBounds(755, 328, 85, 32);
		
		textQuantity4 = new Text(shell, SWT.BORDER);
		textQuantity4.setText("0");
		textQuantity4.setBounds(755, 366, 85, 32);
		
		//Payment Info
		labelPaymentType = new Label(shell, SWT.NONE);
		labelPaymentType.setText("Payment Type");
		labelPaymentType.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.NORMAL));
		labelPaymentType.setBounds(1021, 220, 133, 26);
		
		txtpayment1 = new Text(shell, SWT.BORDER);
		txtpayment1.setBounds(1021, 252, 133, 32);
		
		txtpayment2 = new Text(shell, SWT.BORDER);
		txtpayment2.setBounds(1021, 290, 133, 32);
		
		txtpayment3 = new Text(shell, SWT.BORDER);
		txtpayment3.setBounds(1021, 328, 133, 32);
		
		txtpayment4 = new Text(shell, SWT.BORDER);
		txtpayment4.setBounds(1021, 366, 133, 32);
	
		
		//Sales Info
		txtTotal = new Text(shell, SWT.BORDER);
		txtTotal.setBounds(1021, 566, 133, 26);
		
		txtSubtotal = new Text(shell, SWT.BORDER);
		txtSubtotal.setBounds(1021, 505, 133, 23);
		
		txtTax = new Text(shell, SWT.BORDER);
		txtTax.setBounds(1076, 534, 78, 26);
		
		Label lblSubtotal = new Label(shell, SWT.NONE);
		lblSubtotal.setFont(SWTResourceManager.getFont("Times New Roman", 11, SWT.NORMAL));
		lblSubtotal.setBounds(945, 508, 70, 20);
		lblSubtotal.setText("SubTotal");
		
		Label lblTax = new Label(shell, SWT.NONE);
		lblTax.setFont(SWTResourceManager.getFont("Times New Roman", 11, SWT.NORMAL));
		lblTax.setBounds(1040, 540, 26, 20);
		lblTax.setText("Tax");
		
		Label lblTotal = new Label(shell, SWT.NONE);
		lblTotal.setFont(SWTResourceManager.getFont("Times New Roman", 11, SWT.NORMAL));
		lblTotal.setBounds(971, 569, 44, 20);
		lblTotal.setText("Total");
						
		// Generate Invoice Button
		Button btnGenerateInvoice = new Button(shell, SWT.BORDER | SWT.CENTER);
		btnGenerateInvoice.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.BOLD));
		btnGenerateInvoice.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM-dd-yyyy");  
				LocalDateTime now = LocalDateTime.now();
				LocalDateTime yesterday = now.minus(1, ChronoUnit.DAYS);
				
				lblStatus.setText("Status: " + "Working.....");
				
				String CustomerName, EagleID, PhoneNum, Email;
				CustomerName = txtCustomerName.getText(); EagleID = txtEmail.getText(); 
				PhoneNum = txtPhoneNum.getText();         Email = txtEmail.getText();
				
				String upc1, upc2, upc3, upc4;
				upc1 = txtUPC.getText();   upc2 = textUPC2.getText();
				upc3 = textUPC3.getText(); upc4 = textUPC4.getText();
				
				String Description, Description2, Description3, Description4;
				Description = textDescription1.getText(); Description2 = textDescription2.getText(); 
				Description3 = textDescription3.getText(); Description4 = textDescription4.getText();
				
				String Quantity1, Quantity2, Quantity3, Quantity4;
				Quantity1 = textQuantity1.getText(); Quantity2 = textQuantity2.getText();
				Quantity3 = textQuantity3.getText(); Quantity4 = textQuantity4.getText();
				
				String Price, Price2, Price3, Price4;
				Price = textPrice1.getText(); Price2 = textPrice2.getText();
				Price3 = textPrice3.getText(); Price4 = textPrice4.getText();
				
				String SerialNum1, SerialNum2, SerialNum3, SerialNum4;
				SerialNum1 = textSerial1.getText(); SerialNum2 = textSerial2.getText();
				SerialNum3 = textSerial3.getText(); SerialNum4 = textSerial4.getText();
				
				String Payment1, Payment2, Payment3, Payment4;
				Payment1 = txtpayment1.getText(); Payment2 = txtpayment2.getText();
				Payment3 = txtpayment3.getText(); Payment4 = txtpayment4.getText();
				
				@SuppressWarnings("unused")
				String date, yesterday1;
				date = dtf.format(now);
				System.out.println(date);
				yesterday1 = dtf.format(yesterday);
				
				String xlfileDictName = "";
				JFileChooser fileChooser = new JFileChooser();
		        fileChooser.setDialogTitle("Open the file"); //name for chooser
		        FileFilter filter = new FileNameExtensionFilter("Files", ".xlsx"); //filter to show only that
		        fileChooser.setAcceptAllFileFilterUsed(true); //to show or not all other files
		        fileChooser.addChoosableFileFilter(filter);
		        fileChooser.setSelectedFile(new File(xlfileDictName)); //when you want to show the name of file into the chooser
		        fileChooser.setVisible(true);
		        int result = fileChooser.showOpenDialog(fileChooser);
		        if (result == JFileChooser.APPROVE_OPTION) {
		            xlfileDictName = fileChooser.getSelectedFile().getAbsolutePath();
		        } else {
		            return;
		        }
				
		        System.out.println(xlfileDictName);
		        
				try {
					XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(xlfileDictName));
					
					Object[][] UnitInfo = {
			                {CustomerName,PhoneNum,Email,EagleID,upc1, Description,Quantity1,SerialNum1, Price, Payment1},
			                {"","","","",upc2, Description2, Quantity2, SerialNum2, Price2, Payment2},
			                {"","","","",upc3, Description3, Quantity3, SerialNum3, Price3, Payment3},
			                {"","","","",upc4, Description4, Quantity4, SerialNum4, Price4, Payment4},
			        };
					
					XSSFSheet sheet = workbook.getSheet(date);
					int Test = workbook.getSheetIndex(date);
					System.out.println(Test);
					if (Test != -1) {
						sheet = workbook.getSheet(date);
						} else {
							workbook.createSheet(date);
						}
					System.out.println("Sheet Checking Done");
					int rowCount = sheet.getLastRowNum();
					
					System.out.println(rowCount);
					
		            for (Object[] Data : UnitInfo) {
		            	Row row = sheet.createRow(++rowCount);
		 
		                int columnCount = -1;
		                 
		                for (Object field : Data) {
		                   Cell cell = row.createCell(++columnCount);
		                   sheet.autoSizeColumn(columnCount);
		                    if (field instanceof String) {
		                        cell.setCellValue((String) field);
		                    } else if (field instanceof Integer) {
		                        cell.setCellValue((Integer) field);
		                    }
		                }
		                System.out.println("Info Finished");
		            }
					System.out.println("Writing File");
					
					FileOutputStream outputStream = new FileOutputStream(xlfileDictName);
					workbook.write(outputStream);
		            workbook.close();
		            outputStream.close();
		            
		            System.out.println("Done!");
				} catch (IOException | EncryptedDocumentException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	            lblStatus.setText("Status: " + "Done!");
				}
		});
		btnGenerateInvoice.setBounds(232, 423, 224, 44);
		btnGenerateInvoice.setText("Create Invoice");
		
		//Receipt Button
		Button btnReceipt = new Button(shell, SWT.BORDER | SWT.CENTER);
		btnReceipt.addSelectionListener(new SelectionAdapter() {
			@SuppressWarnings({ "unused", "resource" })
			@Override
			public void widgetSelected(SelectionEvent e) {
				DateTimeFormatter dtf = DateTimeFormatter.ofPattern("MM-dd-yyyy");  
				LocalDateTime now = LocalDateTime.now();
				LocalDateTime yesterday = now.minus(1, ChronoUnit.DAYS);
				
				lblStatus.setText("Status: " + "Working.....");
				
				String CustomerName, EagleID, PhoneNum, Email;
				CustomerName = txtCustomerName.getText(); EagleID = txtEmail.getText(); 
				PhoneNum = txtPhoneNum.getText();         Email = txtEmail.getText();
				
				String upc1, upc2, upc3, upc4;
				upc1 = txtUPC.getText();   upc2 = textUPC2.getText();
				upc3 = textUPC3.getText(); upc4 = textUPC4.getText();
				
				String Description, Description2, Description3, Description4;
				Description = textDescription1.getText(); Description2 = textDescription2.getText(); 
				Description3 = textDescription3.getText(); Description4 = textDescription4.getText();
				
				String Quantity1, Quantity2, Quantity3, Quantity4;
				Quantity1 = textQuantity1.getText(); Quantity2 = textQuantity2.getText();
				Quantity3 = textQuantity3.getText(); Quantity4 = textQuantity4.getText();
				
				double dQuantity1, dQuantity2, dQuantity3, dQuantity4;
				dQuantity1 = Double.parseDouble(Quantity1); dQuantity2 = Double.parseDouble(Quantity2);
				dQuantity3 = Double.parseDouble(Quantity3); dQuantity4 = Double.parseDouble(Quantity4);
				
				String Price, Price2, Price3, Price4;
				Price = textPrice1.getText(); Price2 = textPrice2.getText();
				Price3 = textPrice3.getText(); Price4 = textPrice4.getText();
				
				 double SubPrice1 = Double.parseDouble(Price), SubPrice2 = Double.parseDouble(Price2), 
						 SubPrice3 = Double.parseDouble(Price3), SubPrice4 = Double.parseDouble(Price);
				
				String SerialNum1, SerialNum2, SerialNum3, SerialNum4;
				SerialNum1 = textSerial1.getText(); SerialNum2 = textSerial2.getText();
				SerialNum3 = textSerial3.getText(); SerialNum4 = textSerial4.getText();
				
				String Payment1, Payment2, Payment3, Payment4;
				Payment1 = txtpayment1.getText(); Payment2 = txtpayment2.getText();
				Payment3 = txtpayment3.getText(); Payment4 = txtpayment4.getText();
				
				String date, yesterday1;
				date = dtf.format(now);
				System.out.println(date);
				yesterday1 = dtf.format(yesterday);
				
				String fileDictName = "";
				JFileChooser fileChooser = new JFileChooser();
				fileChooser.setDialogTitle("Save As"); //name for chooser
				FileFilter filter = new FileNameExtensionFilter("doc", ".docx"); //filter to show only that
				fileChooser.setAcceptAllFileFilterUsed(true); //to show or not all other files
				fileChooser.addChoosableFileFilter(filter);
				fileChooser.setSelectedFile(new File(fileDictName + CustomerName + "-" + date + ".docx")); //when you want to show the name of file into the chooser
				fileChooser.setVisible(true);
				int result = fileChooser.showOpenDialog(fileChooser);
				if (result == JFileChooser.APPROVE_OPTION) {
				    fileDictName = fileChooser.getSelectedFile().getAbsolutePath();
				} else {
				    return;
				}
				System.out.println(fileDictName);

				try {
					XWPFDocument doc = new XWPFDocument(OPCPackage.open("input.docx"));
					XWPFDocument docout = new XWPFDocument();
					for (XWPFParagraph p : doc.getParagraphs()) {
					    List<XWPFRun> runs = p.getRuns();
					    if (runs != null) {
					        for (XWPFRun r : runs) {
					            String text = r.getText(0);
					            
					            if (text != null && text.contains("<name>")) {
					                text = text.replace("<name>", CustomerName + " - Email:" + Email + " - Phone:" + PhoneNum);
					                r.setText(text, 0);
					            }
										       				            
					            if (text != null && text.contains("<date>")) {
					                text = text.replace("<date>", date);
					                r.setText(text, 0);
					            }			            
					            //Product 1
					            if (text != null && text.contains("<product1>")) {
					            	if(upc1 == "") {
					            		text = text.replace("<product1>", "No Product");
						                r.setText(text, 0);
					            	}else {
					                text = text.replace("<product1>", upc1 + " - " + Description + " @ $" + dQuantity1 * SubPrice1 + " " + Payment1);
					                r.setText(text, 0);
					            	}
					            }
					            if (text != null && text.contains("<serial1>")) {
					                text = text.replace("<serial1>", SerialNum1);
					                r.setText(text, 0);
					            }
					            //Product 2
					            if (text != null && text.contains("<product2>")) {
					            	if(upc2 == "") {
					            		text = text.replace("<product2>", "No Product");
						                r.setText(text, 0);
					            	}else {
					                text = text.replace("<product2>", upc2 + " - " + Description2 + " @ $" + dQuantity2 * SubPrice2 + " " + Payment2);
					                r.setText(text, 0);
					            	}
					            }
					            if (text != null && text.contains("<serial2>")) {
					                text = text.replace("<serial2>", SerialNum2);
					                r.setText(text, 0);
					            }
					            //Product 3
					            if (text != null && text.contains("<product3>")) {
					            	if(upc3 == "") {
					            		text = text.replace("<product3>", "No Product");
						                r.setText(text, 0);
					            	}else {
					                text = text.replace("<product3>", upc3 + " - " + Description3 + " @ $" + dQuantity3 * SubPrice3 + " " + Payment3);
					                r.setText(text, 0);
					            	}
					            }
					            if (text != null && text.contains("<serial3>")) {
						            text = text.replace("<serial3>", SerialNum3);
						            r.setText(text, 0);
						            }
					            //Product 4
					            if (text != null && text.contains("<product4>")) {
					            	if(upc4 == "") {
					            		text = text.replace("<product4>", "No Product");
						                r.setText(text, 0);
					            	}else {
					                text = text.replace("<product4>", upc4 + " - " + Description4 + " @ $" + dQuantity4 * SubPrice4 + " - " + Payment4);
					                r.setText(text, 0);
					            	}
					            }
					            if (text != null && text.contains("<serial4>")) {
						            text = text.replace("<serial4>", SerialNum4);
						            r.setText(text, 0);
						            }
					        }
					    }
					}
					System.out.println(fileDictName);
					FileOutputStream outputStream = new FileOutputStream(fileDictName);
					doc.write(outputStream);
					outputStream.close();
					
				} catch (IOException | org.apache.poi.openxml4j.exceptions.InvalidFormatException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} 
				
				lblStatus.setText("Status: " + "Done!");
				
			}
		});
		btnReceipt.setText("Receipt");
		btnReceipt.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.BOLD));
		btnReceipt.setBounds(554, 423, 224, 44);
		
		// Total Button
		Button btnTotal = new Button(shell, SWT.BORDER | SWT.CENTER);
		btnTotal.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				String Unit1, Unit2, Unit3, Unit4;
				Unit1 = textPrice1.getText(); Unit2 = textPrice2.getText();
				Unit3 = textPrice3.getText(); Unit4 = textPrice4.getText();
				
				String Quantity1, Quantity2, Quantity3, Quantity4;
				Quantity1 = textQuantity1.getText(); Quantity2 = textQuantity2.getText();
				Quantity3 = textQuantity3.getText(); Quantity4 = textQuantity4.getText();
				
				System.out.println(Unit1);
				
			    double Price1 = Double.parseDouble(Unit1), Price2 = Double.parseDouble(Unit2), 
			    		Price3 = Double.parseDouble(Unit3), Price4 = Double.parseDouble(Unit4);
			    
			    double dQuantity1, dQuantity2, dQuantity3, dQuantity4;
				dQuantity1 = Double.parseDouble(Quantity1); dQuantity2 = Double.parseDouble(Quantity2);
				dQuantity3 = Double.parseDouble(Quantity3); dQuantity4 = Double.parseDouble(Quantity4);
				
			    if(dQuantity1 != 0) {
			    	Price1 = Price1 * dQuantity1; 
			    }
			    
			    if(dQuantity2 != 0) {
			    	Price2 = Price2 * dQuantity2; 
			    }
			    
			    if(dQuantity3 != 0) {
			    	Price3 = Price3 * dQuantity3; 
			    }
			    
			    if(dQuantity4 != 0) {
			    	Price4 = Price4 * dQuantity4; 
			    }
			    
			    double
				SubTotal = Price1 + Price2 + Price3 + Price4,
				rSubTotal = Math.round(SubTotal * 100.0) / 100.0,
				Total = SubTotal * 1.08,
				rTotal = Math.round(Total * 100.0) / 100.0,
				TaxTotal = Total - SubTotal,
			    rTaxTotal = Math.round(TaxTotal * 100.0) / 100.0;
				
				String strSub = String.valueOf(rSubTotal), strTax = String.valueOf(rTaxTotal), strTotal = String.valueOf(rTotal);
			 
				txtSubtotal.setText(strSub);
				txtTax.setText(strTax);
				txtTotal.setText(strTotal);
			}
		});
		btnTotal.setFont(SWTResourceManager.getFont("Times New Roman", 14, SWT.BOLD));
		btnTotal.setBounds(930, 423, 224, 44);
		btnTotal.setText("Total");

	
	}
}
