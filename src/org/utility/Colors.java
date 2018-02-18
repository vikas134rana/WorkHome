package org.utility;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public class Colors {

	private static Map<String, Short> colorList = new HashMap<>();

	static {
		colorList.put("YELLOW", IndexedColors.YELLOW.getIndex());
		colorList.put("BLACK", IndexedColors.BLACK.getIndex());
		colorList.put("BROWN", IndexedColors.BROWN.getIndex());
		colorList.put("OLIVE_GREEN", IndexedColors.OLIVE_GREEN.getIndex());
		colorList.put("DARK_GREEN", IndexedColors.DARK_GREEN.getIndex());
		colorList.put("DARK_TEAL", IndexedColors.DARK_TEAL.getIndex());
		colorList.put("DARK_BLUE", IndexedColors.DARK_BLUE.getIndex());
		colorList.put("INDIGO", IndexedColors.INDIGO.getIndex());
		colorList.put("GREY_80_PERCENT", IndexedColors.GREY_80_PERCENT.getIndex());
		colorList.put("ORANGE", IndexedColors.ORANGE.getIndex());
		colorList.put("DARK_YELLOW", IndexedColors.DARK_YELLOW.getIndex());
		colorList.put("GREEN", IndexedColors.GREEN.getIndex());
		colorList.put("TEAL", IndexedColors.TEAL.getIndex());
		colorList.put("BLUE", IndexedColors.BLUE.getIndex());
		colorList.put("BLUE_GREY", IndexedColors.BLUE_GREY.getIndex());
		colorList.put("GREY_50_PERCENT", IndexedColors.GREY_50_PERCENT.getIndex());
		colorList.put("RED", IndexedColors.RED.getIndex());
		colorList.put("LIGHT_ORANGE", IndexedColors.LIGHT_ORANGE.getIndex());
		colorList.put("LIME", IndexedColors.LIME.getIndex());
		colorList.put("SEA_GREEN", IndexedColors.SEA_GREEN.getIndex());
		colorList.put("AQUA", IndexedColors.AQUA.getIndex());
		colorList.put("LIGHT_BLUE", IndexedColors.LIGHT_BLUE.getIndex());
		colorList.put("VIOLET", IndexedColors.VIOLET.getIndex());
		colorList.put("GREY_40_PERCENT", IndexedColors.GREY_40_PERCENT.getIndex());
		colorList.put("PINK", IndexedColors.PINK.getIndex());
		colorList.put("GOLD", IndexedColors.GOLD.getIndex());
		colorList.put("YELLOW", IndexedColors.YELLOW.getIndex());
		colorList.put("BRIGHT_GREEN", IndexedColors.BRIGHT_GREEN.getIndex());
		colorList.put("TURQUOISE", IndexedColors.TURQUOISE.getIndex());
		colorList.put("DARK_RED", IndexedColors.DARK_RED.getIndex());
		colorList.put("SKY_BLUE", IndexedColors.SKY_BLUE.getIndex());
		colorList.put("PLUM", IndexedColors.PLUM.getIndex());
		colorList.put("GREY_25_PERCENT", IndexedColors.GREY_25_PERCENT.getIndex());
		colorList.put("ROSE", IndexedColors.ROSE.getIndex());
		colorList.put("LIGHT_YELLOW", IndexedColors.LIGHT_YELLOW.getIndex());
		colorList.put("LIGHT_GREEN", IndexedColors.LIGHT_GREEN.getIndex());
		colorList.put("LIGHT_TURQUOISE", IndexedColors.LIGHT_TURQUOISE.getIndex());
		colorList.put("PALE_BLUE", IndexedColors.PALE_BLUE.getIndex());
		colorList.put("LAVENDER", IndexedColors.LAVENDER.getIndex());
		colorList.put("WHITE", IndexedColors.WHITE.getIndex());
		colorList.put("CORNFLOWER_BLUE", IndexedColors.CORNFLOWER_BLUE.getIndex());
		colorList.put("LEMON_CHIFFON", IndexedColors.LEMON_CHIFFON.getIndex());
		colorList.put("MAROON", IndexedColors.MAROON.getIndex());
		colorList.put("ORCHID", IndexedColors.ORCHID.getIndex());
		colorList.put("CORAL", IndexedColors.CORAL.getIndex());
		colorList.put("ROYAL_BLUE", IndexedColors.ROYAL_BLUE.getIndex());
		colorList.put("LIGHT_CORNFLOWER_BLUE", IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
		colorList.put("TAN", IndexedColors.TAN.getIndex());
	}

	public Short getColorIndex(String code) {
		Short index = colorList.get(code.toUpperCase());
		if (index == null)
			index = colorList.get("BLACK");
		return index;
	}

	public void setcolorName(String code, Short name) {
		colorList.put(code, name);
	}

	public static void main(String[] args) {
		System.out.println(new ColorsOld().getColorIndex("WHITE"));
	}

	public void setTextColor(Workbook wb, Cell cell, String color) {
		CellStyle cellStyle = wb.createCellStyle();
		Font font = wb.createFont();
		cellStyle.setFont(font);
		font.setColor(getColorIndex(color.toUpperCase().trim()));
		cell.setCellStyle(cellStyle);
	}
	
	public void setCellColor(Workbook wb, Cell cell, String color) {
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setFillForegroundColor(getColorIndex(color.toUpperCase().trim()));
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(cellStyle);
	}
}
