package org.utility;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelObject {

	private String excelPath;
	private String excelReference;
	private Workbook wb;
	private Sheet sh;

	public String getExcelPath() {
		return excelPath;
	}

	public void setExcelPath(String excelPath) {
		this.excelPath = excelPath;
	}

	public String getExcelReference() {
		return excelReference;
	}

	public void setExcelReference(String excelReference) {
		this.excelReference = excelReference;
	}

	public Workbook getWb() {
		return wb;
	}

	public void setWb(Workbook wb) {
		this.wb = wb;
	}

	public Sheet getSh() {
		return sh;
	}

	public void setSh(Sheet sh) {
		this.sh = sh;
	}

	// name with extension
	public String getName() {
		return FilenameUtils.getName(this.excelPath);
	}

	// name without extension
	public String getBaseName() {
		return FilenameUtils.getBaseName(this.excelPath);
	}

	public String getExtension() {
		return FilenameUtils.getExtension(this.excelPath);
	}

}
