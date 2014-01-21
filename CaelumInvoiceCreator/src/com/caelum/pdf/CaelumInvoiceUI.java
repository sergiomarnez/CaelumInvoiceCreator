package com.caelum.pdf;

import java.io.File;
import java.io.IOException;
import javax.swing.UIManager;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

import com.caelum.pdf.PDFReportCreator;
import com.caelum.pdf.PdfInfo;
import com.itextpdf.text.DocumentException;
import com.tools.excel.ReadExcel;

import org.eclipse.wb.swt.SWTResourceManager;

public class CaelumInvoiceUI{

	private Shell shlCealumInvoiceCreator;
	private Text lineNumber;
	private Text pathname;
	private Text pathnameD;
	private Label lblDestinoDeReporte;
	private Button btnElegirExcel;
	private Button createButton;
	private Button btnExplorar;

	/**
	 * Launch the application.
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			UIManager.put("FileChooser.readOnly", Boolean.TRUE);
			CaelumInvoiceUI window = new CaelumInvoiceUI();
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
		shlCealumInvoiceCreator.open();
		shlCealumInvoiceCreator.layout();
		while (!shlCealumInvoiceCreator.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}

	/**
	 * Create contents of the window.
	 */
	protected void createContents() {		
		
		shlCealumInvoiceCreator = new Shell();
		shlCealumInvoiceCreator.setSize(531, 300);
		shlCealumInvoiceCreator.setText("Cealum Invoice Creator");
		shlCealumInvoiceCreator.setLayout(null);
		
		pathname = new Text(shlCealumInvoiceCreator, SWT.BORDER);
		pathname.setBounds(35, 55, 393, 21);
		pathname.setEnabled(false);
		
		btnElegirExcel = new Button(shlCealumInvoiceCreator, SWT.NONE);
		btnElegirExcel.setBounds(434, 53, 55, 25);
		btnElegirExcel.getShell().forceActive();
		btnElegirExcel.setText("Explorar");
		btnElegirExcel.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				
				FileDialog dlg = new FileDialog(btnElegirExcel.getShell(),  SWT.OPEN  );
				dlg.setText("Elige el documento Excel de Facturas");
				dlg.setFilterExtensions(new String[] { "*.xlsx", "*.xls" }); // Windows
                // wild
                // cards
				String path = dlg.open();
				if (path == null){ return;
				
				}else{
					pathname.setText(dlg.getFilterPath()+"\\"+dlg.getFileName());
				}
			}
		});
		
		pathnameD = new Text(shlCealumInvoiceCreator, SWT.BORDER);
		pathnameD.setBounds(35, 181, 393, 21);
		pathnameD.setEnabled(false);
		
		btnExplorar = new Button(shlCealumInvoiceCreator, SWT.NONE);
		btnExplorar.setBounds(434, 179, 55, 25);
		btnExplorar.setText("Explorar");
		btnExplorar.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				FileDialog dlg = new FileDialog(btnElegirExcel.getShell(),  SWT.SAVE  );
				dlg.setText("Elige el fichero de destino PDF");
				dlg.setFilterExtensions(new String[] { "*.pdf" }); // Windows
                // wild
                // cards
				String path = dlg.open();
				if (path == null){
					return;				
				}else{
					pathnameD.setText(dlg.getFilterPath()+"\\"+dlg.getFileName());
				}
				
			}
		});
		

		createButton = new Button(shlCealumInvoiceCreator, SWT.NONE);
		createButton.setBounds(192, 218, 116, 25);
		createButton.setText("Crear Factura PDF");
		createButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				if(checkFullData()){
					PdfInfo p = new PdfInfo();
					p = migrateExcelLineToPdfInfo(Integer.valueOf(lineNumber.getText())-1,"pathExcel");
					try {
						createReport(p,pathnameD.getText());
						if ((new File(pathnameD.getText())).exists()) {
							Process pr = Runtime.getRuntime().exec("rundll32 url.dll,FileProtocolHandler "+pathnameD.getText());
							pr.waitFor();
						} else {

							System.out.println("File is not exists");
						}
						System.out.println("Done");

					} catch (DocumentException | IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} catch (InterruptedException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
				}
			}
		});
		

		Label lblNewLabel = new Label(shlCealumInvoiceCreator, SWT.NONE);
		lblNewLabel.setBounds(35, 88, 66, 15);
		lblNewLabel.setText("Linea Excel:");		
		
		lineNumber = new Text(shlCealumInvoiceCreator, SWT.BORDER);
		lineNumber.setBounds(111, 85, 27, 21);

		Label lblExcel = new Label(shlCealumInvoiceCreator, SWT.NONE);
		lblExcel.setBounds(35, 34, 150, 15);
		lblExcel.setText("Excel facturas Caelum:");

		Label lblej = new Label(shlCealumInvoiceCreator, SWT.NONE);
		lblej.setBounds(144, 88, 55, 15);
		lblej.setText("(ej: 20)");

		lblDestinoDeReporte = new Label(shlCealumInvoiceCreator, SWT.NONE);
		lblDestinoDeReporte.setForeground(SWTResourceManager.getColor(SWT.COLOR_BLACK));
		lblDestinoDeReporte.setBounds(35, 160, 164, 15);
		lblDestinoDeReporte.setText("Destino de la factura PDF:");		

		Label label = new Label(shlCealumInvoiceCreator, SWT.SEPARATOR | SWT.HORIZONTAL);
		label.setBounds(27, 133, 462, 2);

	}


	public void createReport(PdfInfo p,String file) throws DocumentException, IOException {

		PDFReportCreator report = new PDFReportCreator();
		report.modifyPdf(p,file);

	}

	public PdfInfo migrateExcelLineToPdfInfo(Integer lineNumber, String excelFileName){
		PdfInfo p = new PdfInfo();

		ReadExcel reader = new ReadExcel();
		try {
			reader.read(p,lineNumber,pathname.getText());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return p;		
	}

	public boolean checkFullData(){

		if(pathname.getText()==null || "".equals(pathname.getText())){
			printError("Factura Excel");
			return false;
		}
		
		if(lineNumber.getText()==null || "".equals(lineNumber.getText())){
			printError("Número de línea");
			return false;
		}

		if(pathnameD.getText()==null || "".equals(pathnameD.getText())){
			printError("Factura PDF");
			return false;
		}
		
		return true;		
	}
	
	public void printError(String field){
		// Message
	    MessageBox messageDialog = new MessageBox(shlCealumInvoiceCreator, SWT.ERROR);
	    messageDialog.setText("Faltan campos por rellenar o están mal completados");
	    messageDialog.setMessage("Por favor introduzca correcamente los datos en el campo "+field);
	    int returnCode = messageDialog.open();
	}
}
