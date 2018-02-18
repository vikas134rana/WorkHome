package org.utility;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.crestech.opkey.plugin.communication.contracts.functionresult.FunctionResult;
import com.crestech.opkey.plugin.communication.contracts.functionresult.Result;

public class ThreadManager implements Runnable {

	private String excelReference;
	private String sheetName;
	private Sheet sh;
	private String output;
	private String delimeter;

	public ThreadManager(String excelReference, String sheetName, String delimeter) {
		this.excelReference = excelReference;
		this.sheetName = sheetName;
		this.delimeter = delimeter;
	}

	public String getFullExcelText() {
		return this.output;
	}

	public Sheet getSheet() {
		return this.sh;
	}

	@Override
	public void run() {
		setExcelWorkbookAndSheet(this.excelReference, this.sheetName);
		this.output = getFullExcelValues(this.excelReference, this.sheetName, this.delimeter);
	}

	public FunctionResult setExcelWorkbookAndSheet(String excelPath, String sheetName) {
		Workbook wb;
		try {
			wb = ExcelPlugin.excelMap.get(excelPath).excelWorkbook;
		} catch (Exception e) {
			return Result.FAIL().setOutput(false).setMessage("Excel Reference <" + excelPath + "> is not opened").make();
		}
		// if (wb == null)
		// return Result.FAIL().setOutput(false).setMessage("Excel File <" +
		// excelPath + "> is not opened").make();
		try {
			int sheetIndex = wb.getSheetIndex(sheetName);
			System.out.println("sheetIndex executed");
			this.sh = (Sheet) wb.getSheetAt(sheetIndex);
			System.out.println("sh executed");
		} catch (Exception e) {
			return Result.FAIL().setOutput(false).setMessage("Sheet <" + sheetName + "> is not present in Excel File <" + excelPath + ">").make();
		}
		return Result.PASS().setOutput(true).make();
	}

	public String getFullExcelValues(String excelPath, String sheetName, String delimeter) {
		int totalRows = getExcelTotalRowCount();
		String excelValues = "";
		for (int i = 1; i <= totalRows; i++) {
			String rowValues = getExcelRowValues(i, delimeter);
			System.out.println(rowValues);
			excelValues = excelValues + rowValues + delimeter;
		}
		return excelValues;
	}

	public int getExcelTotalRowCount() {
		int totalRows = this.sh.getLastRowNum() + 1;
		return totalRows;
	}

	public String getExcelRowValues(int row, String delimeter) {
		if (delimeter.isEmpty())
			delimeter = ";";
		String cellData = "";
		Row row1 = this.sh.getRow(row - 1);
		if (row1 == null)
			row1 = this.sh.createRow(row - 1);
		String rowData = "";
		System.out.println(row1);
		int rowLength = row1.getLastCellNum();
		// System.out.println("rowLength: " + rowLength);
		for (int i = 1; i <= rowLength; i++) {
			cellData = getExcelCellValue(row, i);
			rowData = rowData + cellData + delimeter;
		}
		if (!rowData.isEmpty())
			rowData = rowData.substring(0, rowData.length() - delimeter.length());
		return rowData;
	}

	public String getExcelCellValue(int row, int col) {
		row--;
		col--;
		String cellData = "";
		try {
			cellData = this.sh.getRow(row).getCell(col).getStringCellValue();
		} catch (Exception e) {
			try {
				cellData = String.valueOf((this.sh.getRow(row).getCell(col).getNumericCellValue()));
				if (cellData.indexOf(".") == cellData.lastIndexOf(".")) {
					try {
						if (Integer.parseInt(cellData.split("\\.")[1]) == 0) {
							cellData = cellData.split("\\.")[0];
						}
					} catch (Exception ex) {
						// TODO: handle exception
					}
				}

			} catch (Exception ex) {
				cellData = "";
			}
		}
		return cellData;
	}

}
