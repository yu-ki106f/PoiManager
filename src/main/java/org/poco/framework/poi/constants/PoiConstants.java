package org.poco.framework.poi.constants;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;

public class PoiConstants {
	
	public enum FLColor {
		BLACK
			(HSSFColor.BLACK.class),
		BROWN
			(HSSFColor.BROWN.class),
		OLIVE_GREEN
			(HSSFColor.OLIVE_GREEN.class),
		DARK_GREEN
			(HSSFColor.DARK_GREEN.class),
		DARK_TEAL
			(HSSFColor.DARK_TEAL.class),
		DARK_BLUE
			(HSSFColor.DARK_BLUE.class),
		INDIGO
			(HSSFColor.INDIGO.class),
		GREY_80_PERCENT
			(HSSFColor.GREY_80_PERCENT.class),
		ORANGE
			(HSSFColor.ORANGE.class),
		DARK_YELLOW
			(HSSFColor.DARK_YELLOW.class),
		GREEN
			(HSSFColor.GREEN.class),
		TEAL
			(HSSFColor.TEAL.class),
		BLUE
			(HSSFColor.BLUE.class),
		BLUE_GREY
			(HSSFColor.BLUE_GREY.class),
		GREY_50_PERCENT
			(HSSFColor.GREY_50_PERCENT.class),
		RED
			(HSSFColor.RED.class),
		LIGHT_ORANGE
			(HSSFColor.LIGHT_ORANGE.class),
		LIME
			(HSSFColor.LIME.class),
		SEA_GREEN
			(HSSFColor.SEA_GREEN.class),
		AQUA
			(HSSFColor.AQUA.class),
		LIGHT_BLUE
			(HSSFColor.LIGHT_BLUE.class),
		VIOLET
			(HSSFColor.VIOLET.class),
		GREY_40_PERCENT
			(HSSFColor.GREY_40_PERCENT.class),
		PINK
			(HSSFColor.PINK.class),
		GOLD
			(HSSFColor.GOLD.class),
		YELLOW
			(HSSFColor.YELLOW.class),
		BRIGHT_GREEN
			(HSSFColor.BRIGHT_GREEN.class),
		TURQUOISE
			(HSSFColor.TURQUOISE.class),
		DARK_RED
			(HSSFColor.DARK_RED.class),
		SKY_BLUE
			(HSSFColor.SKY_BLUE.class),
		PLUM
			(HSSFColor.PLUM.class),
		GREY_25_PERCENT
			(HSSFColor.GREY_25_PERCENT.class),
		ROSE
			(HSSFColor.ROSE.class),
		LIGHT_YELLOW
			(HSSFColor.LIGHT_YELLOW.class),
		LIGHT_GREEN
			(HSSFColor.LIGHT_GREEN.class),
		LIGHT_TURQUOISE
			(HSSFColor.LIGHT_TURQUOISE.class),
		PALE_BLUE
			(HSSFColor.PALE_BLUE.class),
		LAVENDER
			(HSSFColor.LAVENDER.class),
		WHITE
			(HSSFColor.WHITE.class),
		CORNFLOWER_BLUE
			(HSSFColor.CORNFLOWER_BLUE.class),
		LEMON_CHIFFON
			(HSSFColor.LEMON_CHIFFON.class),
		MAROON
			(HSSFColor.MAROON.class),
		ORCHID
			(HSSFColor.ORCHID.class),
		CORAL
			(HSSFColor.CORAL.class),
		ROYAL_BLUE
			(HSSFColor.ROYAL_BLUE.class),
		LIGHT_CORNFLOWER_BLUE
			(HSSFColor.LIGHT_CORNFLOWER_BLUE.class),
		TAN
			(HSSFColor.TAN.class);

		/**
		 * クラスを保持
		 */
		private Class<? extends HSSFColor> clazz;
		
		/**
		 * コンストラクタ
		 * @param value
		 */
		private FLColor(Class<? extends HSSFColor> value) {
			this.clazz = value;
		}
		
		/**
		 * index値を返す
		 * @return
		 */
		public short value() {
			try {
				return Short.parseShort(clazz.getField("index").get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
		
		/**
		 * index値を返す
		 * @return
		 */
		public short index() {
			return value();
		}
	}

	public enum Border {
		BORDER_NONE,
		BORDER_THIN,
		BORDER_MEDIUM,
		BORDER_DASHED,
		BORDER_DOTTED,
		BORDER_THICK,
		BORDER_DOUBLE,
		BORDER_HAIR,
		BORDER_MEDIUM_DASHED,
		BORDER_DASH_DOT,
		BORDER_MEDIUM_DASH_DOT,
		BORDER_DASH_DOT_DOT,
		BORDER_MEDIUM_DASH_DOT_DOT,
		BORDER_SLANTED_DASH_DOT;
		
		
		public short value() {
			try {
				return Short.parseShort(CellStyle.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
	
	public enum Align {
		ALIGN_GENERAL,
		ALIGN_LEFT,
		ALIGN_CENTER,
		ALIGN_RIGHT,
		ALIGN_FILL,
		ALIGN_JUSTIFY,
		ALIGN_CENTER_SELECTION;
		
		
		public short value() {
			try {
				return Short.parseShort(CellStyle.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
	
	public enum VAlign {
		VERTICAL_TOP,
		VERTICAL_CENTER,
		VERTICAL_BOTTOM ,
		VERTICAL_JUSTIFY;

		public short value() {
			try {
				return Short.parseShort(CellStyle.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}		
		
	public enum FillPattern {
		NO_FILL,
		SOLID_FOREGROUND,
		FINE_DOTS,
		ALT_BARS,
		SPARSE_DOTS,
		THICK_HORZ_BANDS,
		THICK_VERT_BANDS,
		THICK_BACKWARD_DIAG,
		THICK_FORWARD_DIAG,
		BIG_SPOTS,
		BRICKS,
		THIN_HORZ_BANDS,
		THIN_VERT_BANDS,
		THIN_BACKWARD_DIAG,
		THIN_FORWARD_DIAG,
		SQUARES,
		DIAMONDS;

		public short value() {
			try {
				return Short.parseShort(CellStyle.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
	
	public enum BorderPosition {
		top,
		left,
		right,
		bottom,
		all;
	}
	
	public enum BoldWeight {
		BOLDWEIGHT_NORMAL,
		BOLDWEIGHT_BOLD;
		  
		public short value() {
			try {
				return Short.parseShort(Font.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
	
	public enum FontOffset {
		SS_NONE,
		SS_SUPER,
		SS_SUB;
		
		public short value() {
			try {
				return Short.parseShort(Font.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
	
	public enum FontUnderline {
		U_NONE,
		U_SINGLE,
		U_DOUBLE,
		U_SINGLE_ACCOUNTING,
		U_DOUBLE_ACCOUNTING;
		
		public byte value() {
			try {
				return Byte.parseByte(Font.class.getField(name()).get(null).toString() );
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
	
	public enum AnchorType {
	    MOVE_AND_RESIZE,
	    MOVE_DONT_RESIZE,
	    DONT_MOVE_AND_RESIZE;
	    
	    public int value() {
			try {
				return Integer.parseInt(ClientAnchor.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
	    }
		
	}
	
	public enum CellType {
		CELL_TYPE_NUMERIC,
		CELL_TYPE_STRING,
		CELL_TYPE_FORMULA,
		CELL_TYPE_BLANK,
		CELL_TYPE_BOOLEAN,
		CELL_TYPE_ERROR;
		
		public int value() {
			try {
				return Integer.parseInt(Cell.class.getField(name()).get(null).toString());
			} catch (NumberFormatException e) {
				e.printStackTrace();
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			} catch (NoSuchFieldException e) {
				e.printStackTrace();
			}
			return -1;
		}
	}
}
