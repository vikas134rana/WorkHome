package org.utility;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Color;

public class ColorsOld {

		
	private static Map<String, Short> colorList = new HashMap<>();
	
		
	
	static {
		colorList.put("YELLOW", HSSFColor.YELLOW.index);
		colorList.put("BLACK", HSSFColor.BLACK.index);
		colorList.put("BROWN", HSSFColor.BROWN.index);
		colorList.put("OLIVE_GREEN", HSSFColor.OLIVE_GREEN.index);
		colorList.put("DARK_GREEN", HSSFColor.DARK_GREEN.index);
		colorList.put("DARK_TEAL", HSSFColor.DARK_TEAL.index);
		colorList.put("DARK_BLUE", HSSFColor.DARK_BLUE.index);
		colorList.put("INDIGO", HSSFColor.INDIGO.index);
		colorList.put("GREY_80_PERCENT", HSSFColor.GREY_80_PERCENT.index);
		colorList.put("ORANGE", HSSFColor.ORANGE.index);
		colorList.put("DARK_YELLOW", HSSFColor.DARK_GREEN.index);
		colorList.put("GREEN", HSSFColor.GREEN.index);
		colorList.put("TEAL", HSSFColor.TEAL.index);
		colorList.put("BLUE", HSSFColor.BLUE.index);
		colorList.put("BLUE_GREY", HSSFColor.BLUE_GREY.index);
		colorList.put("GREY_50_PERCENT", HSSFColor.GREY_50_PERCENT.index);
		colorList.put("RED", HSSFColor.RED.index);
		colorList.put("LIGHT_ORANGE", HSSFColor.LIGHT_ORANGE.index);
		colorList.put("LIME", HSSFColor.LIME.index);
		colorList.put("SEA_GREEN", HSSFColor.SEA_GREEN.index);
		colorList.put("AQUA", HSSFColor.AQUA.index);
		colorList.put("LIGHT_BLUE", HSSFColor.LIGHT_BLUE.index);
		colorList.put("VIOLET", HSSFColor.VIOLET.index);
		colorList.put("GREY_40_PERCENT", HSSFColor.GREY_40_PERCENT.index);
		colorList.put("PINK", HSSFColor.PINK.index);
		colorList.put("GOLD", HSSFColor.GOLD.index);
		colorList.put("YELLOW", HSSFColor.YELLOW.index);
		colorList.put("BRIGHT_GREEN", HSSFColor.BRIGHT_GREEN.index);
		colorList.put("TURQUOISE", HSSFColor.TURQUOISE.index);
		colorList.put("DARK_RED", HSSFColor.DARK_RED.index);
		colorList.put("SKY_BLUE", HSSFColor.SKY_BLUE.index);
		colorList.put("PLUM", HSSFColor.PLUM.index);
		colorList.put("GREY_25_PERCENT", HSSFColor.GREY_25_PERCENT.index);
		colorList.put("ROSE", HSSFColor.ROSE.index);
		colorList.put("LIGHT_YELLOW", HSSFColor.LIGHT_YELLOW.index);
		colorList.put("LIGHT_GREEN", HSSFColor.LIGHT_GREEN.index);
		colorList.put("LIGHT_TURQUOISE", HSSFColor.LIGHT_TURQUOISE.index);
		colorList.put("PALE_BLUE", HSSFColor.PALE_BLUE.index);
		colorList.put("LAVENDER", HSSFColor.LAVENDER.index);
		colorList.put("WHITE", HSSFColor.WHITE.index);
		colorList.put("CORNFLOWER_BLUE", HSSFColor.CORNFLOWER_BLUE.index);
		colorList.put("LEMON_CHIFFON", HSSFColor.LEMON_CHIFFON.index);
		colorList.put("MAROON", HSSFColor.MAROON.index);
		colorList.put("ORCHID", HSSFColor.ORCHID.index);
		colorList.put("CORAL", HSSFColor.CORAL.index);
		colorList.put("ROYAL_BLUE", HSSFColor.ROYAL_BLUE.index);
		colorList.put("LIGHT_CORNFLOWER_BLUE", HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
		colorList.put("TAN", HSSFColor.TAN.index);
	}

	public Short getColorIndex(String code) {
		Short index = colorList.get(code.toUpperCase());
		if(index == null)
			index = colorList.get("BLACK");
		return index;
	}

	public void setcolorName(String code, Short name) {
		colorList.put(code, name);
	}

	public static void main(String args[]) {
		System.out.println(new ColorsOld().getColorIndex("WHITE"));
	}

}
