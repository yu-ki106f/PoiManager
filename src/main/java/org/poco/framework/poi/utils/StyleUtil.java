package org.poco.framework.poi.utils;

import org.poco.framework.poi.constants.PoiConstants.Align;
import org.poco.framework.poi.constants.PoiConstants.BoldWeight;
import org.poco.framework.poi.constants.PoiConstants.Border;
import org.poco.framework.poi.constants.PoiConstants.FLColor;
import org.poco.framework.poi.constants.PoiConstants.FillPattern;
import org.poco.framework.poi.constants.PoiConstants.FontOffset;
import org.poco.framework.poi.constants.PoiConstants.FontUnderline;
import org.poco.framework.poi.constants.PoiConstants.VAlign;

public class StyleUtil {
	
	public static FLColor getColor(short index) {
		for (FLColor item : FLColor.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}

	public static Border getBorder(short index) {
		for (Border item : Border.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}
	
	public static Align getAlign(short index) {
		for (Align item : Align.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}

	public static VAlign getVAlign(short index) {
		for (VAlign item : VAlign.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}

	public static FillPattern getFillPattern(short index) {
		for (FillPattern item : FillPattern.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}

	public static BoldWeight getBoldWeight(short index) {
		for (BoldWeight item : BoldWeight.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}
	
	public static FontUnderline getFontUnderline(byte index) {
		for (FontUnderline item : FontUnderline.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}
	
	public static FontOffset getFontOffset(short index) {
		for (FontOffset item : FontOffset.values() ) {
			if (index == item.value()) {
				return item;
			}
		}
		return null;
	}
}
