package com.tools.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.caelum.pdf.PdfInfo;

public class ReadExcel {

	public PdfInfo read(PdfInfo p, Integer lineNumber,String filename) throws IOException  {
		File inputWorkbook = new File(filename);

		FileInputStream fis = new FileInputStream(inputWorkbook);
		XSSFWorkbook wb = new XSSFWorkbook(fis);

		XSSFSheet sh = wb.getSheetAt(1);

		System.out.println(sh.getLastRowNum());
		System.out.println("Name: "+sh.getSheetName()); 
		Row row = sh.getRow(lineNumber);


		//Nº orden
		if(row.getCell(0) != null){
			p.setNorden(row.getCell(0).getStringCellValue());	  
		}

		//Fecha
		if(row.getCell(1) != null){
			p.setFecha(row.getCell(1).getDateCellValue());	
		}
		//nº factura
		if(row.getCell(2) != null){
			p.setNfactura(row.getCell(2).getStringCellValue());	 
		}
		//Nif
		if(row.getCell(3) != null){
			p.setNif(row.getCell(3).getStringCellValue());	
		}
		//Nombre
		if(row.getCell(4) != null){
			p.setNombre(row.getCell(4).getStringCellValue());	  
		}
		//Concepto
		if(row.getCell(5) != null){
			p.setConcepto(row.getCell(5).getStringCellValue());	 
		}
		//Iva
		if(row.getCell(6) != null){
			p.setIva(row.getCell(6).getNumericCellValue());
		}
		//Retenido
		if(row.getCell(7) != null){
			p.setRetencion(row.getCell(7).getNumericCellValue());
		}
		//Deducible
		if(row.getCell(8) != null){
			p.setDeducible(row.getCell(8).getNumericCellValue());
		}
		//No deducible
		if(row.getCell(9) != null){
			p.setNoDeducible(row.getCell(9).getNumericCellValue());
		}
		//Total
		if(row.getCell(10) != null){
			p.setTotal(row.getCell(10).getNumericCellValue());
		}

		System.out.println("Contenido de la fila: "+p.toString());
		return p;
	}
} 
