package org.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.crestech.opkey.plugin.ResultCodes;
import com.crestech.opkey.plugin.communication.contracts.functionresult.FunctionResult;
import com.crestech.opkey.plugin.communication.contracts.functionresult.Result;
import com.crestech.opkey.plugin.utility.exceptionHandler.ArgumentDataMissingException;

public class Utils {

	public static Sheet sh = null;
	public static Workbook wb = null;
	public static String cellDelimeter = ";";
	public static String excelPath = null;

	/*
	 * public String getExcelCellValue(int row, int col) { row--; col--; String
	 * cellData = "";
	 * 
	 * try { cellData = sh.getRow(row).getCell(col).getStringCellValue(); } catch
	 * (Exception e) { try { cellData =
	 * String.valueOf((sh.getRow(row).getCell(col).getNumericCellValue())).split
	 * ("\\.")[0]; } catch (Exception ex) { cellData = ""; } } return cellData; }
	 */

	public String getExcelCellValue2(int row, int col) {
		row--;
		col--;
		String cellData = "";

		try {
			Cell cell = sh.getRow(row).getCell(col);
			DataFormatter df = new DataFormatter();
			cellData = df.formatCellValue(cell);
			// cellData = cell.getStringCellValue();
			System.out.println("ABC");
		} catch (Exception e) {
			try {
				System.out.println("XYZ");
				cellData = String.valueOf((sh.getRow(row).getCell(col).getNumericCellValue()));
				if (cellData.contains("E") || cellData.contains("e")) {
					BigDecimal bd = new BigDecimal(cellData);
					long lonVal = bd.longValue();
					cellData = String.valueOf(lonVal);
					return cellData;
				}
				if (cellData.indexOf(".") == cellData.lastIndexOf(".")) {
					try {
						if (Integer.parseInt(cellData.split("\\.")[1]) == 0) {
							cellData = cellData.split("\\.")[0];
						}
					} catch (Exception ex) {
					}
				}

			} catch (Exception ex) {
				cellData = "";
			}
		}
		return cellData;
	}

	public Cell getExcelCell(int row, int col) {
		row--;
		col--;

		Row row1 = sh.getRow(row);
		if (row1 == null)
			row1 = sh.createRow(row);
		Cell cell = row1.getCell(col);
		if (cell == null)
			cell = row1.createCell(col);

		/*
		 * Cell cell = null; cell = sh.getRow(row).getCell(col);
		 */
		return cell;
	}

	public String getExcelRowValues1(int row, String delimeter) {
		if (delimeter.isEmpty())
			delimeter = ";";
		String cellData = "";
		Row row1 = sh.getRow(row - 1);
		if (row1 == null)
			row1 = sh.createRow(row - 1);
		String rowData = "";
		// System.out.println(row1);
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

	public String getExcelColumnValues(int col, String delimeter) {
		if (delimeter.isEmpty())
			delimeter = ";";
		String columnData = "", celldata = "";
		int totalRows = getExcelTotalRowCount();
		System.out.println("totalRows: " + totalRows);
		for (int i = 1; i <= totalRows; i++) {
			celldata = getExcelCellValue(i, col);
			System.out.println(i + "," + col + "  : " + "<" + celldata + ">");
			columnData = columnData + celldata + delimeter;
		}
		columnData = columnData.substring(0, columnData.length() - delimeter.length());
		System.out.println("Column Data: " + columnData);
		return columnData;
	}

	public int getExcelTotalRowCount() {
		int totalRows = sh.getLastRowNum() + 1;
		return totalRows;
	}

	/*
	 * public static int lastColumnInExcel(String excelPath, String sheetName) {
	 * setExcelWorkbookAndSheet(excelPath, sheetName); CTSheetDimension dimension =
	 * sh.getCTWorksheet().getDimension(); String sheetDimensions =
	 * dimension.getRef(); System.out.println(sheetDimensions); char ch =
	 * sheetDimensions.charAt(sheetDimensions.indexOf(":") + 1); int col = ((int) ch
	 * - 64); System.out.println(col); return col; }
	 */

	public FunctionResult setExcelWorkbookAndSheet(String excelPath, String sheetName) {
		sh = null;
		wb = null;
		try {
			wb = ExcelPlugin.excelMap.get(excelPath).excelWorkbook;
			Utils.excelPath = ExcelPlugin.excelMap.get(excelPath).excelPath;
		} catch (Exception e) {
			return Result.FAIL().setOutput(false).setMessage("Excel Reference <" + excelPath + "> is not opened")
					.make();
		}
		// if (wb == null)
		// return Result.FAIL().setOutput(false).setMessage("Excel File <" +
		// excelPath + "> is not opened").make();
		try {
			int sheetIndex = wb.getSheetIndex(sheetName);
			System.out.println("sheetIndex executed");
			sh = (Sheet) wb.getSheetAt(sheetIndex);
			System.out.println("sh executed");
		} catch (Exception e) {
			return Result.FAIL().setOutput(false)
					.setMessage("Sheet <" + sheetName + "> is not present in Excel File <" + excelPath + ">").make();
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

	public String comparison_failed(String value1, String value2, String text) {
		if (text.isEmpty())
			text = "Value";
		return "1st " + text + " :<" + value1 + "> \n2nd " + text + " :<" + value2 + ">";
	}

	public String varification_failed(String orignal, String expected) {
		return " Orignal Value " + " :<" + orignal + "> \n" + "   Expected Value " + " :<" + expected + ">";
	}

	public String varification_failed(int orignal, int expected) {
		return " Orignal Value " + " :<" + orignal + "> \n" + "   Expected Value " + " :<" + expected + ">";
	}

	public String getDelimiter() {
		return ";";
	}

	public String getFileExtension(String fileName) {
		String extension = "";
		int i = fileName.lastIndexOf('.');
		if (i > 0) {
			extension = fileName.substring(i + 1);
		}
		return extension;
	}

	public FunctionResult setValueToExcel(String excelReference) throws IOException {
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(ExcelPlugin.excelMap.get(excelReference).getExcelPath());
			Utils.wb.write(fos);
			wb.close();
		} catch (Exception e) {
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION)
					.setMessage("Close the Excel File <" + excelReference + "> to modify the data of file")
					.setOutput(false).make();
		}
		// fos.flush();
		fos.close();
		fos = null;
		Utils.wb = null;
		Utils.sh = null;
		new Utils().Shadow_ExcelOpen(ExcelPlugin.excelMap.get(excelReference).getExcelPath(), excelReference);
		return Result.PASS().setOutput(true).make();
	}

	public FunctionResult Shadow_ExcelOpen(String excelPath, String excelReference) {
		Workbook wb = null;
		// XSSFSheet sh;
		try {
			FileInputStream fis = new FileInputStream(new File(excelPath));
			System.out.println("fis executed");
			if (new Utils().getFileExtension(excelPath).equalsIgnoreCase("xlsx")) {
				wb = new XSSFWorkbook(fis);
			} else if (new Utils().getFileExtension(excelPath).equalsIgnoreCase("xls")) {
				wb = new HSSFWorkbook(fis);
			} else
				return Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID)
						.setMessage("Check Specified Excel File path (supported excel file extensions .xlsx and .xls)")
						.setOutput(false).make();
			System.out.println("wb executed");
			// openedExcelFile = excelPath;
		} catch (Exception e) {
			return Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID)
					.setMessage("Check Specified Excel File path (supported excel file extensions .xlsx and .xls)")
					.setOutput(false).make();
		}

		/*
		 * try { int sheetIndex = wb.getSheetIndex(sheetName);
		 * System.out.println("sheetIndex executed"); sh = (XSSFSheet)
		 * wb.getSheetAt(sheetIndex); System.out.println("sh executed"); } catch
		 * (Exception e) { return Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID).
		 * setMessage("Check Specified Sheet Name").setOutput(false).make(); }
		 */
		/*
		 * if (excelMap.size() > 0) { if (excelMap.get(excelReference) == null) return
		 * Result.FAIL(ResultCodes.ERROR_ARGUMENT_DATA_INVALID).
		 * setMessage("Excel Reference <" + excelReference +
		 * "> is already opened").setOutput(false).make(); }
		 */
		System.out.println("******************************************");
		ExcelObjectOld excelObject = new ExcelObjectOld();
		excelObject.setExcelPath(excelPath);
		excelObject.setExcelWorkbook(wb);
		ExcelPlugin.excelMap.put(excelReference, excelObject);
		return Result.PASS().setOutput(true).make();

	}

	public FunctionResult Shadow_ExcelSetCellBackgroundColor(String excelReference, String sheetName, int rowNumber,
			int colNumber, String color) throws IOException {
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false)
					.make();
		Cell cell = new Utils().getExcelCell(rowNumber, colNumber);
		CellStyle cellStyle = Utils.wb.createCellStyle();
		cellStyle.setFillForegroundColor(new ColorsOld().getColorIndex(color));
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		// cellStyle.setBorderTop((short) 1); // single line border
		// cellStyle.setBorderBottom((short) 1); // single line border
		cell.setCellStyle(cellStyle);
		FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return frr;

	}

	public FunctionResult Shadow_ExcelSetCellBackgroundColorForComparision(String excelReference, String sheetName,
			int rowNumber, int colNumber, String color) throws IOException {
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false)
					.make();
		Cell cell = new Utils().getExcelCell(rowNumber, colNumber);
		CellStyle cellStyle = Utils.wb.createCellStyle();
		cellStyle.setFillForegroundColor(new ColorsOld().getColorIndex(color));
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		// cellStyle.setBorderTop((short) 1); // single line border
		// cellStyle.setBorderBottom((short) 1); // single line border
		cell.setCellStyle(cellStyle);
		// FunctionResult frr = new Utils().setValueToExcel(excelReference);
		return Result.PASS().setOutput(true).make();
	}

	void emptyMapReferencesExcept(String excelReference) {
		for (Map.Entry<String, ExcelObjectOld> entry : ExcelPlugin.excelMap.entrySet()) {
			String key = entry.getKey();
			ExcelObjectOld value = entry.getValue();
			System.out.println(key + "/t" + value.excelPath);
			if (entry.getKey().equals(excelReference)) {
				System.out.println("Excel Reference Found");
				continue;
			} else {
				ExcelPlugin.excelMap.get(key).excelWorkbook = null;
			}
		}
	}

	// ######################################################################################################

	// indexOne = true : row and column starts from 1 else 0
	public String getExcelCellValue(int row, int col, boolean indexOne) {
		if (indexOne) {
			row--;
			col--;
		}
		String cellData = "";
		try {
			Cell cell = sh.getRow(row).getCell(col);
			cellData = new DataFormatter().formatCellValue(cell);
		} catch (Exception ex) {
			cellData = "";
		}
		return cellData;
	}
	
	public String getExcelRowValues(int row, String delimeter, boolean indexOne) {
		if(indexOne)
			row--;
		if (delimeter.isEmpty())
			delimeter = ";";
		String cellData = "";
		Row rowObj = sh.getRow(row);
		if (rowObj == null)
			rowObj = sh.createRow(row);
		String rowData = "";
		int rowLength = rowObj.getLastCellNum();
		for (int i = 1; i <= rowLength; i++) {
			cellData = getExcelCellValue(row, i);
			rowData = rowData + cellData + delimeter;
		}
		if (!rowData.isEmpty())
			rowData = rowData.substring(0, rowData.length() - delimeter.length());
		return rowData;
	}
	

	// ######################################################################################################

}
