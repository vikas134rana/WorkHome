package org.test;

import org.apache.poi.hssf.util.HSSFColor;

public class Color {

	static {
		System.out.println("Hssf: " + HSSFColor.BLACK.index);
	}

	public static void main(String[] args) {

		System.out.println("awt: " + java.awt.Color.BLACK);

	}
}
