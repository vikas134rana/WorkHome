package org.test;

import java.awt.Color;
import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.utility.Colors;
import org.utility.Excel;
import org.utility.Utils0;

public class ExcelBasicOperation {

	public static void main(String[] args) throws InvalidFormatException, IOException {

		System.out.println(Color.black.getRGB());
		
		new Excel().openExcel("excel1", "C:\\Users\\Vinay\\Desktop\\Test - Copy.xlsx");
		new Excel().openExcel("excel2", "C:\\Users\\Vinay\\Desktop\\Test.xlsx");

		System.out.println(new Excel().getMap().size());
		
		new Excel().readyExcel("excel1", "Sheet1");
		Sheet sh = new Excel().getExcelObject("excel1").getSh();
		
		new Colors().setTextColor(new Excel().getExcelObject("excel1").getWb(), new Utils0(sh).getCell(1, 1), "green");
		new Colors().setCellColor(new Excel().getExcelObject("excel1").getWb(), new Utils0(sh).getCell(3, 3), "green");
		new Excel().writeExcel(new Excel().getExcelObject("excel1").getWb(), "C:\\Users\\Vinay\\Desktop\\Test - Copy1.xlsx");
		
		System.err.println("*******CELL********");
		System.out.println("@1 " + new Utils0(sh).getCellValue(1, 1));
		System.out.println("@2 " + new Utils0(sh).getCellValue(1, 20));
		System.out.println("@3 " + new Utils0(sh).getCellValue(3, 1));
		System.out.println("@4 " + new Utils0(sh).getCellValue(4, 1));

		System.err.println("*******ROW********");
		System.out.println("@1 " + new Utils0(sh).getRowValues(0, ""));
		System.out.println("@2 " + new Utils0(sh).getRowValues(0, "  "));
		System.out.println("@3 " + new Utils0(sh).getRowValues(0, "@#@"));
		System.out.println("@4 " + new Utils0(sh).getRowValues(2, " "));
		System.out.println("@4 " + new Utils0(sh).getRowValues(9, " ") + ">");

		System.err.println("*******COLUMN********");
		System.out.println("@1 " + new Utils0(sh).getColumnValues(4, ""));
		System.out.println("@2 " + new Utils0(sh).getColumnValues(2, "  "));
		System.out.println("@3 " + new Utils0(sh).getColumnValues(20, "@#@"));

		System.err.println("******LAST_ROW******");
		System.out.println(new Utils0(sh).getLastRowNum(sh));

		System.err.println("******COL_NUM******");
		System.out.println(new Utils0(sh).getColumnNum(1, "This is a sample description"));
		System.out.println(new Utils0(sh).getColumnNum(10, "This is a sample description"));
		System.out.println(new Utils0(sh).getColumnNum(1, "This is a sample don"));

		System.err.println("******ROW_NUM******");
		System.out.println(new Utils0(sh).getRowNum(4, "This is a sample description"));
		System.out.println(new Utils0(sh).getRowNum(10, "This is a sample description"));
		System.out.println(new Utils0(sh).getRowNum(1, "This is a sample don"));

		System.err.println("******COL_NUM(INDEX)******");
		System.out.println(new Utils0(sh).getColumnNum(1, "This is a sample description", 2));
		System.out.println(new Utils0(sh).getColumnNum(1, "some notes", 1));
		System.out.println(new Utils0(sh).getColumnNum(1, "some notes", 2));

		System.err.println("******ROW_NUM(INDEX)******");
		System.out.println(new Utils0(sh).getRowNum(4, "This is a sample description", 2));
		System.out.println(new Utils0(sh).getRowNum(5, "some notes", 1));
		System.out.println(new Utils0(sh).getRowNum(5, "some notes", 2));

		System.err.println("******ROW_COL_NUM(INDEX)******");
		System.out.println(new Utils0(sh).getRowColNum("This is a sample description", 2));
		System.out.println(new Utils0(sh).getRowColNum("some notes", 1));
		System.out.println(new Utils0(sh).getRowColNum("some notes", 2));
		System.out.println(new Utils0(sh).getRowColNum("some notes", 3));

		System.err.println("******SET_ROW_VALUE******");
		System.out.println(new Utils0(sh).setRowValue(3, "A;B;C;D;E;F;G;H;I;J;K;L;M", ";"));
		System.out.println(new Utils0(sh).setRowValue(11, "A1;B1;C1;D1;E1;F1;G1;H1;I1;J1;K1;L1;M1", ";"));

		System.err.println("******SET_COL_VALUE******");
		System.out.println(new Utils0(sh).setColumnValue(3, "A;B;C;D;E;F;G;H;I;J;K;L;M", ";"));
		System.out.println(new Utils0(sh).setColumnValue(15, "A1;B1;C1;D1;E1;F1;G1;H1;I1;J1;K1;L1;M1", ";"));

		System.err.println("******WRITE_EXCEL******");
//		System.out.println(new Utils0(sh).writeExcel(wb, "C:\\Users\\Vinay\\Desktop\\Test.xlsx"));
	}
}
