package org.utility;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.crestech.opkey.plugin.KeywordLibrary;
import com.crestech.opkey.plugin.ResultCodes;
import com.crestech.opkey.plugin.communication.contracts.functionresult.FunctionResult;
import com.crestech.opkey.plugin.communication.contracts.functionresult.Result;
import com.crestech.opkey.plugin.utility.Validations;
import com.crestech.opkey.plugin.utility.exceptionHandler.ArgumentDataMissingException;
import com.crestech.opkey.plugin.utility.keywords.ExcelPlugin;
import com.crestech.opkey.plugin.utility.keywords.Utils;

public class CompareExcel implements KeywordLibrary {
	// static boolean compareSheetStatus = true;
	static String tempExcelFileReference = "tempOpkey";
	static Sheet tempExcelSheet = null;
	static String tempExcelFileName = "ExcelComparison";
	static String previousTempExcelFilePath = "";
	static int cells = 0;
	public static void main(String[] args) {

		try { // get input excel files

			File file1 = new File("C:\\Users\\vikas.rana\\Desktop\\ExcelCompare\\ExcelCompare1.xlsx");
			File file2 = new File("C:\\Users\\vikas.rana\\Desktop\\ExcelCompare\\ExcelCompare2.xlsx");

			FileInputStream excellFile1 = new FileInputStream(new File("C:\\Users\\vikas.rana\\Desktop\\ExcelCompare\\ExcelCompare1.xlsx"));
			FileInputStream excellFile2 = new FileInputStream(new File("C:\\Users\\vikas.rana\\Desktop\\ExcelCompare\\ExcelCompare2.xlsx"));

			String tempExcel = getTempExcelPath();

			try {
				copyFile(file1, new File(tempExcel));
			} catch (Exception e) {
			}

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook1 = new XSSFWorkbook(excellFile1);
			XSSFWorkbook workbook2 = new XSSFWorkbook(excellFile2);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet1 = workbook1.getSheetAt(0);
			XSSFSheet sheet2 = workbook2.getSheetAt(0);

			// Compare sheets
			if (compareTwoSheets(sheet1, sheet2, new File(tempExcel))) {
				System.out.println("\n\nThe two excel sheets are Equal");
			} else {
				System.out.println("\n\nThe two excel sheets are Not Equal");
			}

			// close files
			excellFile1.close();
			excellFile2.close();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public FunctionResult Custom_ExcelCompareSheets(String excelReference1, String sheetName1, String excelReference2, String sheetName2, String tempExcelName)
			throws ArgumentDataMissingException, IOException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		tempExcelFileName = tempExcelName;
		String tempExcelPath = getTempExcelPath();
		previousTempExcelFilePath = tempExcelPath.replace(".xlsx", "99887766.xlsx");
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference1, sheetName1);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		Sheet sheet1 = Utils.sh;
		String excelPathFile1 = Utils.excelPath;
		 System.out.println("excelPathFile1: " + excelPathFile1);
		 System.out.println("tempExcelPath: " + tempExcelPath);
		try {
			copyFile(new File(excelPathFile1), new File(previousTempExcelFilePath));
		} catch (Exception e) {
			e.printStackTrace();
		}
		FunctionResult fr1 = new Utils().setExcelWorkbookAndSheet(excelReference2, sheetName2);
		Sheet sheet2 = Utils.sh;
		if (fr1.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr1.getMessage()).setOutput(false).make();
		boolean status = true;
		try {
			status = compareTwoSheets(sheet1, sheet2, new File(tempExcelPath));
			new Utils().setValueToExcel(tempExcelFileReference);
			removeExcelMapReference(tempExcelFileReference);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		try{
			File file = new File(previousTempExcelFilePath);
			file.delete();
		}catch (Exception e) {
			e.printStackTrace();
		}
		if (status)
			return Result.PASS().setOutput(tempExcelPath).make();
		else
			return Result.FAIL(ResultCodes.ERROR_VERIFICATION_FAILLED).setOutput(tempExcelPath).make();
	}

	// Compare Two Sheets
	public static boolean compareTwoSheets(Sheet sheet1, Sheet sheet2, File file) throws IOException {
		// System.out.println("###############################################################");
		int firstRow1 = sheet1.getFirstRowNum();
		int lastRow1 = sheet1.getLastRowNum();
		// System.out.println("firstRow1: " + firstRow1);
		// System.out.println("LastRow1: " + lastRow1);
		setTempExcelFile(getTempExcelPath(), previousTempExcelFilePath, sheet1.getSheetName());
		boolean equalSheets = true;
		for (int i = firstRow1; i <= lastRow1; i++) {
			rowCount = i;
			// System.out.println("\n\nComparing Row " + i);

			Row row1 = sheet1.getRow(i);
			Row row2 = sheet2.getRow(i);
			if (!compareTwoRows(row1, row2, sheet1)) {
				equalSheets = false;
				// System.out.println("Row " + i + " - Not Equal");
			} else {
				// System.out.println("Row " + i + " - Equal");
			}
		}
		// System.out.println("###############################################################");
		return equalSheets;
	}

	// Compare Two Rows
	static int rowCount = 0;

	public static boolean compareTwoRows(Row row1, Row row2, Sheet sheet1) throws IOException {
		if ((row1 == null) && (row2 == null)) {
			return true;
		} else if ((row1 == null) || (row2 == null)) {
			// new
			// Utils().Shadow_ExcelSetCellBackgroundColorForComparision(tempExcelFileReference,
			// sheet1.getSheetName(), rowCount + 1, 1, "ROSE");
			return false;
		}

		int firstCell1 = row1.getFirstCellNum();
		int lastCell1 = row1.getLastCellNum();
		boolean equalRows = true;

		// Compare all cells in a row
		for (int i = firstCell1; i <= lastCell1; i++) {
			Cell cell1 = row1.getCell(i);
			Cell cell2 = row2.getCell(i);
			System.out.println("Cell: "+(++cells));
			if (!compareTwoCells(cell1, cell2)) {
				equalRows = false;
				// System.err.println(" Cell " + i + " - NOt Equal" + "Cell1: "
				// + cell1 + " & " + "Cell1: " + cell2);
				setRowColValuesFailComp(row1.getRowNum() + 1, i + 1);
				new Utils().Shadow_ExcelSetCellBackgroundColorForComparision(tempExcelFileReference, sheet1.getSheetName(), rowCount + 1, i + 1, "ROSE");

			} else {
				// System.out.println(" Cell " + i + " - Equal" + "Cell1: " +
				// cell1 + " & " + "Cell1: " + cell2);
			}
		}
		return equalRows;
	}

	// Compare Two Cells
	public static boolean compareTwoCells(Cell cell1, Cell cell2) {
		// System.out.println(11111);
		if ((cell1 == null) && (cell2 == null)) {
			return true;
		} else if ((cell1 == null) || (cell2 == null)) {
			return false;
		}

		// System.out.println(44444);
		boolean equalCells = false;
		int type1 = cell1.getCellType();
		int type2 = cell2.getCellType();
		if (type1 == type2) {
			// System.out.println(22222);
			/* if (cell1.getCellStyle().equals(cell2.getCellStyle())) { */
			// System.out.println(33333);
			// Compare cells based on its type
			switch (type1) {
			case HSSFCell.CELL_TYPE_FORMULA:
				if (cell1.getCellFormula().equals(cell2.getCellFormula())) {
					equalCells = true;
				}
				break;
			case HSSFCell.CELL_TYPE_NUMERIC:
				if (cell1.getNumericCellValue() == cell2.getNumericCellValue()) {
					equalCells = true;
				}
				break;
			case HSSFCell.CELL_TYPE_STRING:
				if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
					equalCells = true;
				}
				break;
			case HSSFCell.CELL_TYPE_BLANK:
				if (cell2.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
					equalCells = true;
				}
				break;
			case HSSFCell.CELL_TYPE_BOOLEAN:
				if (cell1.getBooleanCellValue() == cell2.getBooleanCellValue()) {
					equalCells = true;
				}
				break;
			case HSSFCell.CELL_TYPE_ERROR:
				if (cell1.getErrorCellValue() == cell2.getErrorCellValue()) {
					equalCells = true;
				}
				break;
			default:
				System.out.println("CALLED FROM DEFAULT");
				if (cell1.getStringCellValue().equals(cell2.getStringCellValue())) {
					equalCells = true;
				}
			}

		} else {
			return false;
		}
		return equalCells;
	}

	public static void copyFile(File source, File dest) throws IOException {
		FileUtils.copyFile(source, dest);
	}

	public static String getTempDir() {
		return System.getProperty("java.io.tmpdir");
	}

	public static String getTempExcelFileName() {
		return tempExcelFileName + ".xlsx";
	}

	public static String getTempExcelPath() {
		return getTempDir() + getTempExcelFileName();
	}

	public static void setTempExcelFile(String tempExcelPath, String oldTempExcelPath, String sheetName) {
		try {
			new ExcelPlugin().Method_ExcelOpen(oldTempExcelPath, tempExcelFileReference);
		} catch (ArgumentDataMissingException e) {
		}
		Workbook wb = ExcelPlugin.excelMap.get(tempExcelFileReference).excelWorkbook;
		
		System.out.println("Before: ");
		new CompareExcel().printTotalSheets(wb);
		
		wb = new CompareExcel().deleteSheetExcept(sheetName, wb);
		
		System.out.println("After: ");
		new CompareExcel().printTotalSheets(wb);
		Sheet sheet = wb.getSheetAt(0);
		System.out.println("Sheet 0 Name: "+sheet.getSheetName());
		System.out.println("Sheet 5-2 Value: "+sheet.getRow(2).getCell(5).getStringCellValue());
		removeExcelMapReference(tempExcelFileReference);
		
		System.out.println("After: removeExcelMapReference ");
		new CompareExcel().printTotalSheets(wb);
		 sheet = wb.getSheetAt(0);
		System.out.println("Sheet 0 Name: "+sheet.getSheetName());
		System.out.println("Sheet 5-2 Value: "+sheet.getRow(2).getCell(5).getStringCellValue());
		
		try {
			FileOutputStream fos = new FileOutputStream(tempExcelPath);
			wb.write(fos);
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
//		System.exit(0);
		try {
			new ExcelPlugin().Method_ExcelOpen(tempExcelPath, tempExcelFileReference);
		} catch (ArgumentDataMissingException e) {
		}
		tempExcelSheet = Utils.wb.createSheet("ComparisonReport");
		Row tempExcelHeaderRow = tempExcelSheet.createRow(0);
		tempExcelHeaderRow.createCell(0).setCellValue("ROW");
		tempExcelHeaderRow.createCell(1).setCellValue("Column");
	}

	// set row column values for cells that didn't match (compare excel)

	public static void setRowColValuesFailComp(int row, int col) {
		// System.out.println("setRowColValuesFailComp");
		// System.out.println("row: " + row + "\t" + "col: " + col);
		int lastRow = tempExcelSheet.getLastRowNum();
		// System.out.println("lastRow: " + lastRow);
		// try {
		//// System.out.println("Last Row: "+new
		// DataFormatter().formatCellValue(tempExcelSheet.getRow(lastRow).getCell(0)));
		//// System.out.println("Last Col: "+new
		// DataFormatter().formatCellValue(tempExcelSheet.getRow(lastRow).getCell(1)));
		// } catch (Exception e) {
		// }
		Row tempExcelHeaderRow = tempExcelSheet.createRow(lastRow + 1);
		tempExcelHeaderRow.createCell(0).setCellValue(row);
		tempExcelHeaderRow.createCell(1).setCellValue(col);
	}

	static void removeExcelMapReference(String excelReference) {
		ExcelPlugin.excelMap.remove(excelReference);
	}

	public static void printMap() {
		for (String name : ExcelPlugin.excelMap.keySet()) {
			String key = name.toString();
			String value = ExcelPlugin.excelMap.get(name).excelPath;
			System.out.println(key + " " + value);
		}
	}

	public static void createTempFile() {
		try {
			File file = new File(getTempExcelPath());
			if (file.createNewFile()) {
				System.out.println("Temp File is created!");
			} else {
				System.out.println("Temp File already exists.");
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public Workbook deleteSheetExcept(String sheetName, Workbook wb) {
		for (int i = wb.getNumberOfSheets() - 1; i >= 0; i--) {
			// This is a place holder. You will insert your logic here to get
			// the sheets that you want.
			if (!wb.getSheetName(i).contentEquals(sheetName)) {
				// Just remove the sheets that don't match your criteria in the
				// if statement above
				wb.removeSheetAt(i);
			}
		}
		return wb;
	}
	
	void printTotalSheets(Workbook wb){
		try{
			System.out.println("Total Sheets: "+wb.getNumberOfSheets());
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
}