package com.caelum.pdf;

import java.util.Date;

public class PdfInfo {
	
	private String norden;
	private String nfactura;
	private String nif;
	private Date fecha;
	private String nombre;
	private String concepto;
	private Double iva;
	private Double retencion;
	private Double deducible;
	private Double noDeducible;
	private Double total;
	
	
	
	public String getNorden() {
		return norden;
	}
	public void setNorden(String norden) {
		this.norden = norden;
	}
	public String getNfactura() {
		return nfactura;
	}
	public void setNfactura(String nfactura) {
		this.nfactura = nfactura;
	}
	public String getNif() {
		return nif;
	}
	public void setNif(String nif) {
		this.nif = nif;
	}
	public Date getFecha() {
		return fecha;
	}
	public void setFecha(Date fecha) {
		this.fecha = fecha;
	}
	public String getNombre() {
		return nombre;
	}
	public void setNombre(String nombre) {
		this.nombre = nombre;
	}
	public String getConcepto() {
		return concepto;
	}
	public void setConcepto(String concepto) {
		this.concepto = concepto;
	}
	public Double getIva() {
		return iva;
	}
	public void setIva(Double iva) {
		this.iva = iva;
	}
	public Double getRetencion() {
		return retencion;
	}
	public void setRetencion(Double retencion) {
		this.retencion = retencion;
	}
	public Double getDeducible() {
		return deducible;
	}
	public void setDeducible(Double deducible) {
		this.deducible = deducible;
	}
	public Double getTotal() {
		return total;
	}
	public void setTotal(Double total) {
		this.total = total;
	}
	public Double getNoDeducible() {
		return noDeducible;
	}
	public void setNoDeducible(Double noDeducible) {
		this.noDeducible = noDeducible;
	}
	@Override
	public String toString() {
		return "PdfInfo [norden=" + norden + ", nfactura=" + nfactura
				+ ", nif=" + nif + ", fecha=" + fecha + ", nombre=" + nombre
				+ ", concepto=" + concepto + ", iva=" + iva + ", retencion="
				+ retencion + ", deducible=" + deducible + ", noDeducible="
				+ noDeducible + ", total=" + total + "]";
	}
	
	
	
}
