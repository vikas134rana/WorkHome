package org.utility;

import org.apache.poi.ss.usermodel.Workbook;

public class ExcelObjectOld {

	private String excelPath;
	private Workbook excelWorkbook;

	public String getExcelPath() {
		return excelPath;
	}

	public void setExcelPath(String excelPath) {
		this.excelPath = excelPath;
	}

	public Workbook getExcelWorkbook() {
		return excelWorkbook;
	}

	public void setExcelWorkbook(Workbook excelWorkbook) {
		this.excelWorkbook = excelWorkbook;
	}

}
