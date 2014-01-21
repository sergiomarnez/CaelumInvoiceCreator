package com.caelum.pdf;


import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Locale;

import com.itextpdf.text.Chunk;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;

/**
 * First iText example: Hello World.
 */
public class PDFReportCreator {

	/**
	 * Creates a PDF file: hello.pdf
	 * @param    args    no arguments needed
	 */
	public static void main(String[] args)
			throws DocumentException, IOException {
		//new PDFReportCreator().createPdf("ej.pdf");
		new PDFReportCreator().modifyPdf(new PdfInfo(),"");
	}


	public void modifyPdf(PdfInfo p,String fileName) throws DocumentException, IOException{
		FontFactory.register("cordiau.ttf", "my_bold_font");
        Font myBoldFont = FontFactory.getFont("my_bold_font");
		BaseFont cordia = BaseFont.createFont("cordiau.ttf", BaseFont.WINANSI,true);
		BaseFont bf = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
		BaseFont cordiaBold = BaseFont.createFont("cordiaub.ttf", BaseFont.WINANSI,true);
		try {			

			PdfReader pdfReader = new PdfReader("plantilla.pdf");

			PdfStamper pdfStamper = new PdfStamper(pdfReader,new FileOutputStream(fileName));

			SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy");
			SimpleDateFormat formateador = new SimpleDateFormat("dd 'de' MMMM 'de' yyyy", new Locale("es", "ES"));

			System.out.println("Año: "+format.format(p.getFecha()));

			//Numero de orden
			PdfContentByte underContent = pdfStamper.getOverContent(1);
			underContent.beginText();
			underContent.setLineWidth((float) 3.0);
			underContent.setFontAndSize(cordiaBold, 16);    // set font and size
			underContent.setTextMatrix(137, 665);   // set x,y position (0,0 is at the bottom left)
			underContent.showText(p.getNfactura()!=null?p.getNfactura():String.valueOf(Math.random()*10000000).substring(0, 6));  // set text
			underContent.endText();

			//Fecha
			underContent.beginText();
			underContent.setLineWidth((float) 3.0);
			underContent.setCharacterSpacing((float) 1.13);
			underContent.setFontAndSize(cordiaBold, 13);    // set font and size
			underContent.setTextMatrix(118, 625);   // set x,y position (0,0 is at the bottom left)
			underContent.showText(format.format(p.getFecha()));  // set text
			underContent.endText();

			//Concepto
			PdfContentByte concept = pdfStamper.getOverContent(1);
			concept.beginText();
			concept.setFontAndSize(cordia, 12);    // set font and size
			concept.setTextMatrix(208, 665);   // set x,y position (0,0 is at the bottom left)
			concept.showText(p.getConcepto());  // set text
			concept.endText();

			//Precio del concepto
			PdfContentByte price = pdfStamper.getOverContent(1);
			price.beginText();
			price.setFontAndSize(cordia,12);    // set font and size
			price.setTextMatrix(450, 665);   // set x,y position (0,0 is at the bottom left)
			price.showText(String.valueOf(Math.rint(p.getDeducible()*100)/100)+" €");  // set text
			price.endText();
			
			//Iva aplicado
			PdfContentByte ivaPercent = pdfStamper.getOverContent(1);
			ivaPercent.beginText();
			ivaPercent.setFontAndSize(cordia, 13);    // set font and size
			ivaPercent.setTextMatrix(208, 590);   // set x,y position (0,0 is at the bottom left)
			ivaPercent.showText("IVA "+String.valueOf(Math.rint(p.getIva()))+" %");  // set text
			ivaPercent.endText();
			
			//IVA respecto del precio
			PdfContentByte iva = pdfStamper.getOverContent(1);
			iva.beginText();
			iva.setFontAndSize(cordia, 13);    // set font and size
			iva.setTextMatrix(450, 590);   // set x,y position (0,0 is at the bottom left)
			iva.showText(String.valueOf(Math.rint(p.getRetencion()*100)/100)+ "€");  // set text
			iva.endText();
			
			//Texto Total
			PdfContentByte totalText = pdfStamper.getOverContent(1);
			totalText.beginText();
			totalText.setFontAndSize(bf, 10);    // set font and size
			totalText.setTextMatrix(400, 560);   // set x,y position (0,0 is at the bottom left)
			totalText.showText("Total");  // set text
			totalText.endText();
			
			
			//Total
			PdfContentByte total = pdfStamper.getOverContent(1);
			total.beginText();
			total.setFontAndSize(cordiaBold, 14);    // set font and size
			total.setTextMatrix(450, 560);   // set x,y position (0,0 is at the bottom left)
			total.showText(String.valueOf(Math.rint(p.getTotal()*100)/100)+ "€");  // set text
			total.endText();
			
			//Observaciones texto
			PdfContentByte notesText = pdfStamper.getOverContent(1);
			notesText.beginText();
			notesText.setFontAndSize(cordiaBold, 14);    // set font and size
			notesText.setTextMatrix(208, 525);   // set x,y position (0,0 is at the bottom left)
			notesText.showText("Observaciones:");  // set text
			notesText.endText();
			
			//Notas de la factura
			PdfContentByte notes = pdfStamper.getOverContent(1);
			ColumnText ct = new ColumnText(notes);
			
			String text1 = "El siguiente encargo se realizó a día "+formateador.format(p.getFecha())+". "
					+"La cuenta a abonar el importe es la siguiente: \n\nEvo Banco:";
			ct.setSimpleColumn(new Phrase(new Chunk(text1, myBoldFont)),208, 520, 435, 200, 12, Element.ALIGN_LEFT | Element.ALIGN_TOP);
			ct.go();
			
			//Cuenta del banco
			PdfContentByte bankAccount = pdfStamper.getOverContent(1);
			bankAccount.beginText();
			bankAccount.setFontAndSize(cordiaBold, 12);    // set font and size
			bankAccount.setTextMatrix(270, 460);   // set x,y position (0,0 is at the bottom left)
			bankAccount.showText("2080 5815 52 304 0003187");  // set text
			bankAccount.endText();
			
			//Cliente texto
			PdfContentByte clientText = pdfStamper.getOverContent(1);
			clientText.beginText();
			clientText.setFontAndSize(cordiaBold, 14);    // set font and size
			clientText.setTextMatrix(120, 525);   // set x,y position (0,0 is at the bottom left)
			clientText.showText("Cliente:");  // set text
			clientText.endText();
			
			//Notas del cliente
			PdfContentByte clientInfo = pdfStamper.getOverContent(1);
			ColumnText ct2 = new ColumnText(clientInfo);
			Chunk ch = new Chunk(p.getNombre()/*+"\n"+"Direccion\n"+"CIF/NIF\n"+p.getNif()*/, myBoldFont);
			Phrase ph = new Phrase(ch);

			ct2.setSimpleColumn(ph,120, 520, 220, 200, 12, Element.ALIGN_RIGHT | Element.ALIGN_TOP);
			ct2.go();

			pdfStamper.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (DocumentException e) {
			e.printStackTrace();
		}

		System.out.println("Pdf modificado");
	}
	

}

