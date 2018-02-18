package org.utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	private static Map<String, ExcelObject> map = new HashMap<>();
	Logger logger = Logger.getLogger(Excel.class.getName());

	int mapSize() {
		return map.size();
	}

	public Map<String, ExcelObject> getMap() {
		return map;
	}

	public void setMap(Map<String, ExcelObject> map) {
		Excel.map = map;
	}

	public ExcelObject getExcelObject(String key) {
		return map.get(key);
	}

	public void setExcelObject(String key, ExcelObject excelObject) {
		map.put(key, excelObject);
		System.out.println("mapSize: " + mapSize());
	}

	public void removeExcelObject(String key) {
		map.remove(key);
	}

	public void printMap() {
		for (Map.Entry<String, ExcelObject> entry : map.entrySet()) {
			String key = entry.getKey();
			String value = entry.getValue().getExcelPath();
			String msg = key + " " + value;
			logger.fine(msg);
		}
	}

	public void openExcel(String reference, String excelPath) throws IOException, InvalidFormatException {
		if (map.containsKey(reference))
			throw new IllegalArgumentException("Specified excel reference is already opened.");
		if (!("xls".equals(getExtension(excelPath)) || "xlsx".equals(getExtension(excelPath))))
			throw new IllegalArgumentException("Check specified excel path (file extension should be xls or xlsx)");

		ExcelObject excelObject = new ExcelObject();
		excelObject.setExcelPath(excelPath);
		excelObject.setExcelReference(reference);
		setExcelObject(reference, excelObject);
	}

	public void readyExcel(String reference, String sheetName)
			throws FileNotFoundException, IOException, InvalidFormatException {
		Workbook wb = null;
		ExcelObject excelObject = getExcelObject(reference);
		if (excelObject == null)
			throw new IllegalArgumentException("Specified Excel is not opened or already closed");
		String excelPath = excelObject.getExcelPath();
		if ("xls".equals(getExtension(excelPath)))
			wb = new HSSFWorkbook(new FileInputStream(excelPath));
		else if ("xlsx".equals(getExtension(excelPath)))
			wb = new XSSFWorkbook(new File(excelPath));
		excelObject.setWb(wb);
		excelObject.setSh(excelObject.getWb().getSheet(sheetName));
		setExcelObject(reference, excelObject);
	}

	public boolean writeExcel(Workbook wb, String filePath) {
		try {
			FileOutputStream fos = new FileOutputStream(new File(filePath));
			wb.write(fos);
			wb.close();
			fos.close();
		} catch (Exception e) {
			logger.log(Level.SEVERE, e.getMessage(), e);
			return false;
		}
		return true;
	}

	public String getExtension(String path) {
		return FilenameUtils.getExtension(path);
	}

	// set workbook and sheet to null for all entries in map
	public void clearWbSh() {
		for (Map.Entry<String, ExcelObject> entry : map.entrySet()) {
			String key = entry.getKey();
			ExcelObject excelObject = entry.getValue();
			excelObject.setWb(null);
			excelObject.setSh(null);
			setExcelObject(key, excelObject);
		}
	}

}
