package org.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;





import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatPart;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.crestech.opkey.plugin.KeywordLibrary;
import com.crestech.opkey.plugin.ResultCodes;
import com.crestech.opkey.plugin.communication.contracts.functionresult.FunctionResult;
import com.crestech.opkey.plugin.communication.contracts.functionresult.Result;
import com.crestech.opkey.plugin.utility.Validations;
import com.crestech.opkey.plugin.utility.exceptionHandler.ArgumentDataMissingException;

public class ExcelPlugin implements KeywordLibrary {

	static String cellDelimeter = ";";
	// String openedExcelFile;

	public static Map<String, ExcelObjectOld> excelMap = new HashMap<String, ExcelObjectOld>();

	public FunctionResult Method_ExcelOpen(String excelPath, String excelReference) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1);
		return new Utils().Shadow_ExcelOpen(excelPath, excelReference);
	}

	public FunctionResult Method_ExcelCloseAll() throws IOException, ArgumentDataMissingException {
		try {
			excelMap.clear();
			return Result.PASS().setOutput(true).make();
		} catch (Exception e) {

		}
		return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setOutput(true).make();
	}

	public FunctionResult Mehod_ExcelClose(String excelReference) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0);

		try {
			if (excelMap.remove(excelReference) == null)
				return Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID).setOutput(false).setMessage("Specified Excel Reference is not opened or already closed").make();
			else
				excelMap.remove(excelReference);
		} catch (Exception e) {
			return Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID).setOutput(false).setMessage("Specified Excel Reference is not opened or already closed").make();
		}
		return Result.PASS().setOutput(true).make();
	}

	public FunctionResult Method_ExcelRenameSheet(String excelReference, String oldSheetName, String newSheetName) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2);

		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, oldSheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();

		Utils.wb.setSheetName(Utils.wb.getSheetIndex(oldSheetName), newSheetName);
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelGetColumnCount(String excelReference, String sheetName, int row) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String texts[] = new Utils().getExcelRowValues(row, "").split(Pattern.quote(";"));
		return Result.PASS().setOutput(texts.length).make();
	}

	public FunctionResult Method_ExcelGetRowCount(String excelReference, String sheetName, int col) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		// count empty cell from last till a non-empty cell comes
		int countBlankCell = 0;
		String texts[] = new Utils().getExcelColumnValues(col, "").split(Pattern.quote(";"));
		for (int i = texts.length - 1; i >= 0; i--) {
			if (texts[i].equals(""))
				countBlankCell++;
			else
				break;
		}
		return Result.PASS().setOutput(texts.length - countBlankCell).make();
	}

	public FunctionResult Method_ExcelGetCellValue(String excelReference, String sheetName, int row, int col) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String cellData = new Utils().getExcelCellValue(row, col);
		return Result.PASS().setOutput(cellData).make();
	}

	public FunctionResult Method_ExcelGetRowValue(String excelReference, String sheetName, int row, String delimeter) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String rowData = new Utils().getExcelRowValues(row, delimeter);
		return Result.PASS().setOutput(rowData).make();
	}

	public FunctionResult Method_ExcelGetColumnValue(String excelReference, String sheetName, int col, String delimeter) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String columnData = new Utils().getExcelColumnValues(col, delimeter);
		return Result.PASS().setOutput(columnData).make();
	}

	/*
	 * public FunctionResult Method_getExcelValueRowColumnCount(String
	 * excelPath, String sheetName, String text, int index) throws
	 * ArgumentDataMissingException { Validations.checkDataForBlank(0, 1, 2);
	 * FunctionResult fr = Utils.setExcelWorkbookAndSheet(excelPath, sheetName);
	 * if (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); int c = -1; for (int i = 1; i <=
	 * Utils.getExcelTotalRowCount(); i++) { String rowData =
	 * Method_getExcelRowValues(excelPath, sheetName, i, "").getOutput(); if
	 * (rowData.contains(text)) { String cellData[] = rowData.split(";"); for
	 * (int j = 0; j < cellData.length; j++) { String s = cellData[j];
	 * System.out.println(s); if (s.equals(text)) { c++; if (c == index) return
	 * Result.PASS().setOutput((i) + "," + (j + 1)).make(); } } } } return
	 * Result.FAIL(ResultCodes.ERROR_TEXT_NOT_FOUND).setOutput("-1").
	 * setMessage("Text <" + "> is not found with index " + index).make(); }
	 */

	/*
	 * public FunctionResult Method_getExcelValueRowCount(String excelPath,
	 * String sheetName, String text, int index) throws
	 * ArgumentDataMissingException { Validations.checkDataForBlank(0, 1, 2);
	 * FunctionResult fr = Utils.setExcelWorkbookAndSheet(excelPath, sheetName);
	 * if (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); int c = -1; for (int i = 1; i <=
	 * Utils.getExcelTotalRowCount(); i++) { String rowData =
	 * Method_getExcelRowValues(excelPath, sheetName, i, "").getOutput(); if
	 * (rowData.contains(text)) { String cellData[] = rowData.split(";"); for
	 * (int j = 0; j < cellData.length; j++) { String s = cellData[j];
	 * System.out.println(s); if (s.equals(text)) { c++; if (c == index) return
	 * Result.PASS().setOutput(i).make(); } } } } return
	 * Result.FAIL(ResultCodes.ERROR_TEXT_NOT_FOUND).setOutput("-1").
	 * setMessage("Text <" + "> is not found with index " + index).make(); }
	 */
	/*
	 * public FunctionResult Method_getExcelValueColumnCount(String excelPath,
	 * String sheetName, String text, int index) throws
	 * ArgumentDataMissingException { Validations.checkDataForBlank(0, 1, 2);
	 * FunctionResult fr = Utils.setExcelWorkbookAndSheet(excelPath, sheetName);
	 * if (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); int c = -1; for (int i = 1; i <=
	 * Utils.getExcelTotalRowCount(); i++) { String rowData =
	 * Method_getExcelRowValues(excelPath, sheetName, i, "").getOutput(); if
	 * (rowData.contains(text)) { String cellData[] = rowData.split(";"); for
	 * (int j = 0; j < cellData.length; j++) { String s = cellData[j];
	 * System.out.println(s); if (s.equals(text)) { c++; if (c == index) return
	 * Result.PASS().setOutput(j + 1).make(); } } } } return
	 * Result.FAIL(ResultCodes.ERROR_TEXT_NOT_FOUND).setOutput("-1").
	 * setMessage("Text <" + "> is not found with index " + index).make(); }
	 */

	public FunctionResult Method_ExcelClearCellValue(String excelReference, String sheetName, int startingRow, int startingColumn, int endingRow, int endingColumn)
			throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1);
		startingRow--;
		startingColumn--;
		if (endingRow == 0)
			endingRow = startingRow + 1;
		if (endingColumn == 0)
			endingColumn = startingColumn + 1;
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();

		for (int i = startingRow; i < endingRow; i++) {
			Row row = Utils.sh.getRow(i);
			// System.out.println(row);
			if (row == null)
				row = Utils.sh.createRow(i);
			for (int j = startingColumn; j < endingColumn; j++) {
				Cell cell = row.getCell(j);
				if (cell == null)
					cell = row.createCell(j);
				System.out.println(cell);
				row.removeCell(cell);
			}
		}

		// Cell cell = Utils.sh.getRow(row).getCell(col);
		// Utils.sh.getRow(row).removeCell(cell);
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	/*
	 * public FunctionResult Method_copyCellValueFromExcel(String excelPath,
	 * String sheetName, int row, int col) throws ArgumentDataMissingException {
	 * Validations.checkDataForBlank(0, 1, 2, 3); FunctionResult fr =
	 * Utils.setExcelWorkbookAndSheet(excelPath, sheetName); if
	 * (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); tempCellValue =
	 * Utils.getExcelCellValue(row, col); return
	 * Result.PASS().setOutput(true).make(); }
	 * 
	 * public FunctionResult Method_pasteCellValueInExcel(String excelPath,
	 * String sheetName, int row, int col) throws IOException,
	 * ArgumentDataMissingException { Validations.checkDataForBlank(0, 1, 2, 3);
	 * row--; col--; FunctionResult fr =
	 * Utils.setExcelWorkbookAndSheet(excelPath, sheetName); if
	 * (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make();
	 * 
	 * Row row1 = Utils.sh.getRow(row); if (row1 == null) row1 =
	 * Utils.sh.createRow(row);
	 * row1.createCell(col).setCellValue(tempCellValue); try { FileOutputStream
	 * fos = new FileOutputStream(excelPath); Utils.wb.write(fos); } catch
	 * (FileNotFoundException ex) { return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).
	 * setMessage("Close the Excel File <" + excelPath +
	 * "> to modify the data of file").setOutput(false).make(); } return
	 * Result.PASS().setOutput(true).make(); }
	 */

	public FunctionResult Method_ExcelClearRows(String excelReference, String sheetName, int startingRow, int numberOfRows) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		startingRow--;
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		// for (int i = startingRowNumber; i < startingRowNumber + numberOfRows;
		// i++) {
		// System.out.println(i);
		for (int i = startingRow; i < startingRow + numberOfRows; i++) {
			System.out.println("Row: " + i);
			Row row1 = Utils.sh.getRow(i);
			if (row1 == null)
				row1 = Utils.sh.createRow(i);
			System.out.println(row1);
			Utils.sh.removeRow(row1);
		}
		// }
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelClearColumn(String excelReference, String sheetName, int startingColNumber, int numberOfColumns) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		startingColNumber--;
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		System.out.println(Utils.sh.getLastRowNum());
		int lastRow = Utils.sh.getLastRowNum();
		System.out.println(lastRow);
		for (int i = 0; i <= lastRow; i++) {
			Row row1 = Utils.sh.getRow(i);
			if (row1 == null)
				row1 = Utils.sh.createRow(i);
			for (int j = startingColNumber; j < startingColNumber + numberOfColumns; j++) {
				// System.out.println(row1);
				Cell cell = row1.getCell(j);
				if (cell == null)
					cell = row1.createCell(j);
				row1.removeCell(cell);
			}
		}
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	/*
	 * public FunctionResult Method_copyExcelRow(String excelPath, String
	 * sheetName, int row) throws ArgumentDataMissingException {
	 * Validations.checkDataForBlank(0, 1, 2); FunctionResult fr =
	 * Utils.setExcelWorkbookAndSheet(excelPath, sheetName); if
	 * (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); tempRowValues =
	 * Utils.getExcelRowValues(row, ""); return
	 * Result.PASS().setOutput(true).make(); }
	 * 
	 * public FunctionResult Method_pasteExcelRow(String excelPath, String
	 * sheetName, int row) throws IOException, ArgumentDataMissingException {
	 * Validations.checkDataForBlank(0, 1, 2); row--; FunctionResult fr =
	 * Utils.setExcelWorkbookAndSheet(excelPath, sheetName); if
	 * (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); Row row1 = Utils.sh.getRow(row);
	 * if (row1 == null) row1 = Utils.sh.createRow(row); // int rowLength =
	 * row1.getLastCellNum(); String[] texts = tempRowValues.split(";"); for
	 * (String s : texts) System.out.println(s); int textLength = texts.length;
	 * System.out.println("TextLength: " + textLength); for (int i = 0; i <
	 * textLength; i++) { Cell cell = row1.createCell(i);
	 * System.out.println("cell: " + cell); cell.setCellValue(texts[i]); } try {
	 * FileOutputStream fos = new FileOutputStream(excelPath);
	 * Utils.wb.write(fos); } catch (FileNotFoundException ex) { return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).
	 * setMessage("Close the Excel File <" + excelPath +
	 * "> to modify the data of file").setOutput(false).make(); } return
	 * Result.PASS().setOutput(true).make(); }
	 */
	/*
	 * public FunctionResult Method_copyExcelColumn(String excelPath, String
	 * sheetName, int col) throws ArgumentDataMissingException {
	 * Validations.checkDataForBlank(0, 1, 2); FunctionResult fr =
	 * Utils.setExcelWorkbookAndSheet(excelPath, sheetName); if
	 * (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); tempColumnValues =
	 * Utils.getExcelColumnValues(col, ""); return
	 * Result.PASS().setOutput(true).make(); }
	 * 
	 * public FunctionResult Method_pasteExcelColumn(String excelPath, String
	 * sheetName, int col) throws IOException, ArgumentDataMissingException {
	 * Validations.checkDataForBlank(0, 1, 2); col--; FunctionResult fr =
	 * Utils.setExcelWorkbookAndSheet(excelPath, sheetName); if
	 * (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); String[] texts =
	 * tempColumnValues.split(";"); int textLength = texts.length;
	 * System.out.println("TextLength: " + textLength); for (int i = 0; i <
	 * textLength; i++) { Row row1 = Utils.sh.getRow(i); if (row1 == null) row1
	 * = Utils.sh.createRow(i); Cell cell = row1.getCell(col); if (cell == null)
	 * cell = row1.createCell(col); cell.setCellValue(texts[i]); } try {
	 * FileOutputStream fos = new FileOutputStream(excelPath);
	 * Utils.wb.write(fos); } catch (FileNotFoundException ex) { return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).
	 * setMessage("Close the Excel File <" + excelPath +
	 * "> to modify the data of file").setOutput(false).make(); } return
	 * Result.PASS().setOutput(true).make(); }
	 */
	/*
	 * public FunctionResult Method_insertExcelRow(String excelPath, String
	 * sheetName, int row, String text) throws IOException,
	 * ArgumentDataMissingException { Validations.checkDataForBlank(0, 1, 2, 3);
	 * row--; FunctionResult fr = Utils.setExcelWorkbookAndSheet(excelPath,
	 * sheetName); if (fr.getOutput().trim().equalsIgnoreCase("false")) return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.
	 * getMessage()).setOutput(false).make(); Utils.sh.shiftRows(row, row, 1);
	 * Row row1 = Utils.sh.getRow(row); if (row1 == null) row1 =
	 * Utils.sh.createRow(row); // int rowLength = row1.getLastCellNum();
	 * String[] texts = text.split(";"); // for (String s : texts) //
	 * System.out.println(s); int textLength = texts.length;
	 * System.out.println("TextLength: " + textLength); for (int i = 0; i <
	 * textLength; i++) { Cell cell = row1.createCell(i);
	 * System.out.println("cell: " + cell); cell.setCellValue(texts[i]); } try {
	 * FileOutputStream fos = new FileOutputStream(excelPath);
	 * Utils.wb.write(fos); } catch (FileNotFoundException ex) { return
	 * Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).
	 * setMessage("Close the Excel File <" + excelPath +
	 * "> to modify the data of file").setOutput(false).make(); } return
	 * Result.PASS().setOutput(true).make(); }
	 */
	public FunctionResult Method_ExcelCompareCellValues(String excelReference1, String sheetName1, int row1, int col1, String excelReference2, String sheetName2, int row2, int col2)
			throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4, 5, 6, 7);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference1, sheetName1);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String cell1 = new Utils().getExcelCellValue(row1, col1);
		FunctionResult fr1 = new Utils().setExcelWorkbookAndSheet(excelReference2, sheetName2);
		if (fr1.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();

		String cell2 = new Utils().getExcelCellValue(row2, col2);
		if (cell1.equals(cell2))
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().comparison_failed(cell1, cell2, "Cell")).make();
	}

	public FunctionResult Method_ExcelCompareRows(String excelReference1, String sheetName1, int row1, String excelReference2, String sheetName2, int row2, String delimeter)
			throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4, 5);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference1, sheetName1);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String rowValue1 = new Utils().getExcelRowValues(row1, delimeter);
		FunctionResult fr1 = new Utils().setExcelWorkbookAndSheet(excelReference2, sheetName2);
		if (fr1.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr1.getMessage()).setOutput(false).make();

		String rowValue2 = new Utils().getExcelRowValues(row2, delimeter);
		if (rowValue1.equals(rowValue2))
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().comparison_failed(rowValue1, rowValue2, "Row")).make();
	}

	public FunctionResult Method_ExcelCompareColumns(String excelReference1, String sheetName1, int col1, String excelReference2, String sheetName2, int col2, String delimeter)
			throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4, 5);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference1, sheetName1);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String column1 = new Utils().getExcelColumnValues(col1, delimeter);
		FunctionResult fr1 = new Utils().setExcelWorkbookAndSheet(excelReference2, sheetName2);
		if (fr1.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr1.getMessage()).setOutput(false).make();

		String column2 = new Utils().getExcelColumnValues(col2, delimeter);
		if (column1.equals(column2))
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().comparison_failed(column1, column2, "Column")).make();
	}

	public FunctionResult Method_ExcelCompareSheet(String excelReference1, String sheetName1, String excelReference2, String sheetName2, String delimeter)
			throws IOException, ArgumentDataMissingException, InterruptedException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		ThreadManager m1 = new ThreadManager(excelReference1, sheetName1, delimeter);
		Thread th1 = new Thread(m1);

		ThreadManager m2 = new ThreadManager(excelReference2, sheetName2, delimeter);
		Thread th2 = new Thread(m2);

		th1.start();
		th2.start();

		th1.join();
		th2.join();

		// by now, both the threads have completed

		String excel1 = m1.getFullExcelText();
		String excel2 = m2.getFullExcelText();

		System.out.println("excel1: " + excel1);
		System.out.println("excel2: " + excel2);

		if (excel1.equals(excel2))
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().comparison_failed(excel1, excel2, "Excel")).make();
	}
	/*
	 * public FunctionResult InsertExcelColumn(int col, String text) throws
	 * IOException { col--; String[] texts = text.split(";"); int textLength =
	 * texts.length; for (String s : texts) System.out.println(s);
	 * 
	 * // sh.get for (int j = Utils.lastColumnInExcel() - 1; j >= col; j--) {
	 * copyExcelColumn(j + 1); deleteExcelColumn(j + 1); pasteExcelColumn(j +
	 * 2); // deleteCellValueFromExcel(i + 1, col + 1); } String[] texts1 =
	 * text.split(";"); int textLength1 = texts1.length;
	 * System.out.println("TextLength: " + textLength1); for (int i = 0; i <
	 * textLength1; i++) { XSSFRow row1 = sh.getRow(i); if (row1 == null) row1 =
	 * sh.createRow(i); row1.createCell(col).setCellValue(texts[i]); }
	 * FileOutputStream fos = new FileOutputStream(openedExcelFile);
	 * wb.write(fos);
	 * 
	 * return Result.PASS().setOutput(true).make(); }
	 */

	// #####################################################################################################################################

	/*
	 * public static boolean Method_verifyExcelCellValue(String text, int row,
	 * int col) { if (text.equals((row, col))) return true; return false; }
	 */

	public FunctionResult Method_ExcelVerifyCellValue(String excelReference, String sheetName, int row, int col, String expectedText) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String cellData = new Utils().getExcelCellValue(row, col);
		if (cellData.equalsIgnoreCase(expectedText))
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().varification_failed(cellData, expectedText)).make();
	}

	public FunctionResult Method_ExcelVerifyColumnCount(String excelReference, String sheetName, int row, int expectedColumnCount) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		String texts[] = new Utils().getExcelRowValues(row, "").split(Pattern.quote(";"));
		if (texts.length == expectedColumnCount)
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().varification_failed(texts.length, expectedColumnCount)).make();
	}

	public FunctionResult Method_ExcelVerifyRowCount(String excelReference, String sheetName, int col, int expectedRowCount) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		// count empty cell from last till a non-empty cell comes
		int countBlankCell = 0;
		String texts[] = new Utils().getExcelColumnValues(col, "").split(Pattern.quote(";"));
		for (int i = texts.length - 1; i >= 0; i--) {
			if (texts[i].equals(""))
				countBlankCell++;
			else
				break;
		}
		if (texts.length == expectedRowCount)
			return Result.PASS().setOutput(true).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(false).setMessage(new Utils().varification_failed(texts.length, expectedRowCount)).make();
	}

	public FunctionResult Method_ExcelSetCellValue(String excelReference, String sheetName, String text, int row, int col) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		row--;
		col--;
		Row row1 = Utils.sh.getRow(row);
		if (row1 == null)
			row1 = Utils.sh.createRow(row);
		Cell cell = row1.getCell(col);
		if (cell == null)
			cell = row1.createCell(col);
		cell.setCellValue(text);
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelSetRowValue(String excelReference, String sheetName, int row, String text, String delimeter) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		row--;
		if (delimeter.isEmpty())
			delimeter = ";";
		Row row1 = Utils.sh.getRow(row);
		if (row1 == null)
			row1 = Utils.sh.createRow(row);
		// int rowLength = row1.getLastCellNum();
		String[] texts = text.split(Pattern.quote(delimeter));
		for (String s : texts)
			System.out.println(s);
		int textLength = texts.length;
		System.out.println("TextLength: " + textLength);
		for (int i = 0; i < textLength; i++) {
			Cell cell = row1.getCell(i);
			if (cell == null)
				cell = row1.createCell(i);
			System.out.println("cell: " + cell);
			cell.setCellValue(texts[i]);
		}
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelSetColumnValue(String excelReference, String sheetName, int col, String text, String delimeter) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		col--;
		if (delimeter.isEmpty())
			delimeter = ";";

		// int rowLength = row1.getLastCellNum();
		String[] texts = text.split(Pattern.quote(delimeter));
		int textLength = texts.length;
		System.out.println("TextLength: " + textLength);
		for (int i = 0; i < textLength; i++) {
			Row row1 = Utils.sh.getRow(i);
			if (row1 == null)
				row1 = Utils.sh.createRow(i);
			row1.createCell(col).setCellValue(texts[i]);
		}
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelSetCellBackgroundColor(String excelReference, String sheetName, int rowNumber, int colNumber, String color) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		Cell cell = new Utils().getExcelCell(rowNumber, colNumber);
		
		XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
		XSSFFont fontOld = style.getFont();
		short oldColorIndex = fontOld.getXSSFColor().getIndexed();

		short x;
		CellStyle cellStyle = Utils.wb.createCellStyle();
		cellStyle.setFillForegroundColor(x=new ColorsOld().getColorIndex(color));
//		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		System.out.println(x);
		Font font = Utils.wb.createFont();
		cellStyle.setFont(font);
		font.setColor(oldColorIndex);
		cell.setCellStyle(cellStyle);
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelSetCellTextColor(String excelReference, String sheetName, int rowNumber, int colNumber, String color) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		Cell cell = new Utils().getExcelCell(rowNumber, colNumber);
		
		/*XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
		short oldColorIndex = style.getFillBackgroundColor();
		System.out.println(oldColorIndex);*/
		
		CellStyle cellStyle = Utils.wb.createCellStyle();
		Font font = Utils.wb.createFont();
		cellStyle.setFont(font);
		font.setColor(new ColorsOld().getColorIndex(color));

		/*cellStyle.setFillBackgroundColor(oldColorIndex);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);*/
		
		
		
		cell.setCellStyle(cellStyle);
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		
		return frr;
	}

	public FunctionResult Method_ExcelDeleteRow(String excelReference, String sheetName, int rowNumber) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();

		boolean isNull = false;
		if (Utils.sh.getRow(rowNumber) == null) {
			Utils.sh.createRow(rowNumber);
			isNull = true;
		}
		Utils.sh.shiftRows(rowNumber, Utils.sh.getLastRowNum(), -1);
		if (isNull) {
			--rowNumber;
			Utils.sh.removeRow(Utils.sh.getRow(rowNumber));
		}
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelDuplicateSheet(String excelReference, String sourceSheetName, int position, String targetSheetName) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sourceSheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		Workbook workBook = Utils.wb;
		int sheetCount = workBook.getNumberOfSheets();
		int index = -1;
		for (int i = 0; i < sheetCount; i++) {
			System.out.println(workBook.getSheetAt(i).getSheetName());
			if (workBook.getSheetAt(i).getSheetName().equals(sourceSheetName)) {
				index = i;
				break;
			}
		}
		if (index < 0)
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage("Source Sheet not found").setOutput(false).make();
		System.out.println(Utils.wb.getSheetIndex(targetSheetName));
		if (workBook.getSheetIndex(targetSheetName) >= 0)
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage("The excel file already contain sheet with a name <" + targetSheetName + ">").setOutput(false).make();
		else
			workBook.cloneSheet(index);
		try {
			workBook.setSheetName(sheetCount, targetSheetName);
			workBook.setSheetOrder(targetSheetName, position);
		} catch (IllegalArgumentException e) {
			System.out.println("@2");
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage("The excel file already contain sheet with a name <" + targetSheetName + ">").setOutput(false).make();
		} catch (IndexOutOfBoundsException e) {
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage("Index of sheet is too large").setOutput(false).make();
		}
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		/*
		 * Utils.wb = null;
		 * Method_ExcelOpen(excelMap.get(excelReference).getExcelPath(),
		 * excelReference);
		 */
		return frr;
	}

	public FunctionResult Method_ExcelInsertBlankRows(String excelReference, String sourceSheetName, int numberOfRows, int position) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sourceSheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		Sheet sheet = Utils.sh;
		int rows = sheet.getLastRowNum();
		System.out.println("************ " + rows);
		sheet.shiftRows(position, rows, numberOfRows);
		for (int i = 0; i < numberOfRows; i++) {
			sheet.createRow(position + i);
		}
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelInsertBlankColumns(String excelReference, String sourceSheetName, int numberOfRows, int position) throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sourceSheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		Workbook workBook = Utils.wb;
		Sheet sheet = Utils.sh;
		Cell cell = new Utils().getExcelCell(1, position);

		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;
	}

	public FunctionResult Method_ExcelCopyAndPasteCellValue(String excelReference1, String sheetName1, int row1, int col1, String excelReference2, String sheetName2, int row2, int col2)
			throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3, 4, 5, 6, 7);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference1, sheetName1);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();

		String tempCellValue = new Utils().getExcelCellValue(row1, col1);

		Cell cellFirst = Utils.sh.getRow(--row1).getCell(--col1);
		CellStyle cs = cellFirst.getCellStyle();

		System.out.println("tempCellValue: " + tempCellValue);
		FunctionResult fr1 = new Utils().setExcelWorkbookAndSheet(excelReference2, sheetName2);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr1.getMessage()).setOutput(false).make();

		row2--;
		col2--;
		Row row = Utils.sh.getRow(row2);
		if (row == null)
			row = Utils.sh.createRow(row2);
		Cell cell = row.getCell(col2);
		if (cell == null)
			cell = row.createCell(col2);
		cell.setCellValue(tempCellValue);
		cell.setCellStyle(cs);
		FunctionResult frr = new Utils().setValueToExcel(excelReference2);
		return frr;
	}

	public FunctionResult Method_ExcelCreateFile(String filePath, String sheetName) throws IOException, ArgumentDataMissingException {
		Validations.checkDataForBlank(0);
		try {
			Validations.checkDataForBlank(1);
		} catch (Exception ex) {
			sheetName = "Sheet1";
		}
		File file = new File(filePath);
		if (new Utils().getFileExtension(filePath).equals("xlsx")) {
			if (!file.exists())
				file.createNewFile();
			XSSFWorkbook workbook = new XSSFWorkbook();
			workbook.createSheet(sheetName);
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close();
		} else if (new Utils().getFileExtension(filePath).equals("xls")) {
			if (!file.exists())
				file.createNewFile();
			HSSFWorkbook workbook = new HSSFWorkbook();
			workbook.createSheet(sheetName);
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close();
		} else {
			return Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID).setOutput(false).setMessage("Not a valid file path.").make();
		}
		return Result.PASS().setOutput(filePath).make();
	}

	
	
}
