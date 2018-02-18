package org.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Pattern;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Utils0 {

	Sheet sh = null;
	Excel excelMap = new Excel();
	Logger logger = Logger.getLogger(Utils0.class.getName());

	public Utils0(Sheet sh) {
		this.sh = sh;
	}

	public String getCellValue(int row, int col) {
		String cellData = "";
		try {
			Cell cell = getCell(row, col);
			cellData = getCellValue(cell);
		} catch (Exception ex) {
			cellData = "";
		}
		return cellData;
	}

	public String getRowValues(int row, String delimiter) {
		Row rowObj = getRow(row);
		return getRowValues(rowObj, delimiter);
	}

	public Cell getCell(int row, int col) {
		Cell cell;
		cell = getRow(row).getCell(col);
		if (cell == null)
			cell = sh.getRow(row).createCell(col);
		return cell;
	}

	public String getCellValue(Cell cell) {
		return new DataFormatter().formatCellValue(cell);
	}

	public Row getRow(int row) {
		Row rowObj;
		rowObj = sh.getRow(row);
		if (rowObj == null)
			rowObj = sh.createRow(row);
		return rowObj;
	}

	// Get row values separated by given delimiter(by default delimiter is ; )
	public String getRowValues(Row rowObj, String delimiter) {
		StringBuilder rowValues = new StringBuilder();
		String rowValuesFinal = "";
		if (delimiter.isEmpty())
			delimiter = ";";
		for (Cell c : rowObj) {
			rowValues.append(getCellValue(c) + delimiter);
		}
		rowValuesFinal = rowValues.toString();
		if (!rowValuesFinal.isEmpty())
			rowValuesFinal = rowValuesFinal.substring(0, rowValuesFinal.length() - delimiter.length());
		return rowValuesFinal;
	}

	public String getColumnValues(int col, String delimiter) {
		String columnValuesFinal = "";
		StringBuilder columnValues = new StringBuilder();
		if (delimiter.isEmpty())
			delimiter = ";";
		for (Row r : sh) {
			Cell cell = getCell(r.getRowNum(), col);
			String cellvalue = getCellValue(cell);
			columnValues.append(cellvalue + delimiter);
		}
		if (!columnValuesFinal.isEmpty())
			columnValuesFinal = columnValuesFinal.substring(0, columnValuesFinal.length() - delimiter.length());
		return columnValuesFinal;
	}

	// unusual behavior (Last row is sometime greater than actual. May because of
	// some space at lower cell)
	public int getLastRowNum(Sheet sh) {
		return sh.getLastRowNum();
	}

	public int getColumnNum(int row, String value, int index) {
		int countIndex = 0;
		if (index == 0)
			index = 1;
		int colNum = -1;
		Row rowObj = getRow(row);
		for (Cell c : rowObj) {
			if (getCellValue(c).trim().equalsIgnoreCase(value.trim()) && (++countIndex) == index) {
				colNum = c.getColumnIndex();
			}
		}
		return colNum;
	}

	public int getRowNum(int col, String value, int index) {
		int countIndex = 0;
		if (index == 0)
			index = 1;
		int rowNum = -1;
		for (Row r : sh) {
			Cell cell = getCell(r.getRowNum(), col);
			if (getCellValue(cell).trim().equalsIgnoreCase(value.trim()) && (++countIndex) == index) {
				rowNum = r.getRowNum();
			}
		}
		return rowNum;
	}

	public String getRowColNum(String value, int index) {
		int countIndex = 0;
		if (index == 0)
			index = 1;
		String rowColNum = "";
		for (Row r : sh) {
			for (Cell c : r) {
				Cell cell = getCell(r.getRowNum(), c.getColumnIndex());
				if (getCellValue(cell).trim().equalsIgnoreCase(value.trim()) && (++countIndex) == index) {
					rowColNum = r.getRowNum() + "," + cell.getColumnIndex();
				}
			}
		}
		return rowColNum;
	}

	public boolean setCellValue(int row, int col, String value) {
		try {
			Cell cell = getCell(row, col);
			setCellValue(cell, value);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	public boolean setCellValue(Cell cell, String value) {
		try {
			cell.setCellValue(value);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	public boolean setRowValue(int row, String value, String delimiter) {
		try {
			List<Cell> rowCells = getRowCells(row);
			String[] values = getValues(value, delimiter);
			createRowCells(row, values.length);
			int i = -1;
			for (Cell c : rowCells) {
				setCellValue(c, values[++i]);
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	public boolean setColumnValue(int col, String value, String delimiter) {
		try {
			List<Cell> colCells = getColumnCells(col);
			String[] values = getValues(value, delimiter);
			createColumnCells(col, values.length);
			int i = -1;
			for (Cell c : colCells) {
				setCellValue(c, values[++i]);
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	int getValuesCount(String text, String delimiter) {
		return text.split(Pattern.quote(delimiter)).length;
	}

	String[] getValues(String text, String delimiter) {
		return text.split(Pattern.quote(delimiter));
	}

	boolean createRowCells(int row, int n) {
		try {
			Row rowObj = getRow(row);
			int lastCellNum = rowObj.getLastCellNum();
			if (n > lastCellNum) {
				for (int i = lastCellNum; i < n; i++) {
					rowObj.createCell(i);
				}
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	boolean createColumnCells(int col, int n) {
		try {
			int lastRowNum = sh.getLastRowNum();
			// n = new last row
			if (n > lastRowNum) {
				for (int i = lastRowNum; i < n; i++) {
					Row rowObj = getRow(i);
					rowObj.createCell(col);
				}
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	

	// **************************************************************************************

	public Workbook deleteSheetExcept(String sheetName, Workbook wb) {
		for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
			if (!wb.getSheetName(i).contentEquals(sheetName)) {
				wb.removeSheetAt(i);
			}
		}
		return wb;
	}

	List<Cell> getRowCells(int row) {
		List<Cell> cellList = new ArrayList<>();
		Row rowObj = getRow(row);
		for (Cell c : rowObj) {
			cellList.add(c);
		}
		return cellList;
	}

	List<Cell> getColumnCells(int col) {
		List<Cell> cellList = new ArrayList<>();
		for (Row r : sh) {
			Cell cell = getCell(r.getRowNum(), col);
			cellList.add(cell);
		}
		return cellList;
	}

	public String getExtension(String path) {
		return FilenameUtils.getExtension(path);
	}

	@Deprecated
	public int getColumnNum(int row, String value) {
		int colNum = -1;
		Row rowObj = getRow(row);
		for (Cell c : rowObj) {
			if (getCellValue(c).trim().equalsIgnoreCase(value.trim()))
				colNum = c.getColumnIndex();
		}
		return colNum;
	}

	@Deprecated
	public int getRowNum(int col, String value) {
		int rowNum = -1;
		for (Row r : sh) {
			Cell cell = getCell(r.getRowNum(), col);
			if (getCellValue(cell).trim().equalsIgnoreCase(value.trim()))
				rowNum = r.getRowNum();
		}
		return rowNum;
	}
	// **************************************************************************************
}
