package org.poco.framework.poi.dto;

import java.lang.reflect.Field;

import org.poco.framework.poi.constants.PoiConstants.Align;
import org.poco.framework.poi.constants.PoiConstants.Border;
import org.poco.framework.poi.constants.PoiConstants.Color;
import org.poco.framework.poi.constants.PoiConstants.FillPattern;
import org.poco.framework.poi.constants.PoiConstants.VAlign;
import org.poco.framework.poi.utils.StyleUtil;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class PoiStyleDto extends AbstractDto {

	public Color fillForegroundColor = null;
	public Color fillBackgroundColor = null;
	public FillPattern fillPattern = null;
	
	public Border borderTop = null;
	public Border borderLeft = null;
	public Border borderRight = null;
	public Border borderBottom = null;
	
	//初期値 BLACK
	public Color topBorderColor = Color.BLACK;
	public Color leftBorderColor = Color.BLACK;
	public Color rightBorderColor = Color.BLACK;
	public Color bottomBorderColor = Color.BLACK;
	
	public Align align = null;
	public VAlign valign = null;
	
	public String dataFormat = null;
	/**
	 * フィールドにしないようにprivate + getter, setter
	 */
	private PoiFontDto font = new PoiFontDto();
	
	/**
	 * スタイル
	 */
	private HSSFCellStyle style = null;
	
	/**
	 * fontDtoを取得します。
	 * @return
	 */
	public PoiFontDto getFont() {
		return font;
	}

	/**
	 * fontDtoを設定します。
	 * @param font
	 */
	public void setFont(PoiFontDto font) {
		this.font = font;
	}

	/**
	 * スタイル
	 * @return
	 */
	public HSSFCellStyle getStyle() {
		return style;
	}
	
	/**
	 * コンストラクタ
	 */
	public PoiStyleDto() {}
	
	/**
	 * コンストラクタ<BR/>
	 * 基本スタイルを反映
	 * @param style
	 */
	private PoiStyleDto(HSSFCellStyle style) {

		if (style == null) return;

		this.fillForegroundColor = StyleUtil.getColor(style.getFillForegroundColor());
		this.fillBackgroundColor = StyleUtil.getColor(style.getFillBackgroundColor());
		this.fillPattern = StyleUtil.getFillPattern(style.getFillPattern());
		
		this.borderTop = StyleUtil.getBorder(style.getBorderTop());
		this.borderLeft = StyleUtil.getBorder(style.getBorderLeft());
		this.borderRight = StyleUtil.getBorder(style.getBorderRight());
		this.borderBottom = StyleUtil.getBorder(style.getBorderBottom());

		this.topBorderColor = StyleUtil.getColor(style.getTopBorderColor());
		this.leftBorderColor = StyleUtil.getColor(style.getLeftBorderColor());
		this.rightBorderColor = StyleUtil.getColor(style.getRightBorderColor());
		this.bottomBorderColor = StyleUtil.getColor(style.getBottomBorderColor());
		
		this.align = StyleUtil.getAlign(style.getAlignment());
		this.valign = StyleUtil.getVAlign(style.getVerticalAlignment());
		
		this.dataFormat = style.getDataFormatString();
		
		//本体
		this.style = style;
		//未対応
		/*
		style.getParentStyle();
		style.getUserStyleName();
		style.getHidden();
		*/
	}
	
	/**
	 * コンストラクタ<BR/>
	 * 基本スタイルを反映
	 * @param style
	 */
	public PoiStyleDto(HSSFCellStyle style, HSSFFont font) {
		this(style);
		this.font = new PoiFontDto(font);
	}

	/**
	 * 更新
	 */
	public void update(HSSFCellStyle style,HSSFWorkbook book) {
		if (fillForegroundColor != null) {
			style.setFillForegroundColor(fillForegroundColor.value());
		}
		if (fillBackgroundColor != null) {
			style.setFillBackgroundColor(fillBackgroundColor.value());
		}
		if (fillPattern != null) {
			style.setFillPattern(fillPattern.value());
		}
		if (borderTop != null) {
			style.setBorderTop(borderTop.value());
		}
		if (borderLeft != null) {
			style.setBorderLeft(borderLeft.value());
		}
		if (borderRight != null) {
			style.setBorderRight(borderRight.value());
		}
		if (borderBottom != null) {
			style.setBorderBottom(borderBottom.value());
		}
		if (topBorderColor != null) {
			style.setTopBorderColor(topBorderColor.value());
		}
		if (leftBorderColor != null) {
			style.setLeftBorderColor(leftBorderColor.value());
		}
		if (rightBorderColor != null){
			style.setRightBorderColor(rightBorderColor.value());
		}
		if (bottomBorderColor != null) {
			style.setBottomBorderColor(bottomBorderColor.value());
		}
		if (align != null) {
			style.setAlignment(align.value());
		}
		if (valign != null) {
			style.setVerticalAlignment(valign.value());
		}
		if (dataFormat != null) {
			style.setDataFormat(book.getCreationHelper().createDataFormat().getFormat(dataFormat));
		}
	}
	
	/**
	 * 本クラスに設定をマージする
	 */
	public PoiStyleDto merge(PoiStyleDto dto) {
		Field[] fields = this.getClass().getFields();
		//フィールド一覧取得
		for (Field f : fields ) {
			try {
				//設定あり
				if (f.get(dto) != null) {
					//反映
					f.set(this, f.get(dto));
				}
			}
			catch (IllegalArgumentException e) {}
			catch (IllegalAccessException e) {}
		}
		return this;
	}

//	/**
//	 * キーを返します。
//	 * @return
//	 */
//	public String toKey() {
//		return arrayToString(fillForegroundColor,fillBackgroundColor,fillPattern,
//							borderTop,borderLeft,borderRight,borderBottom,
//							topBorderColor,leftBorderColor,rightBorderColor,bottomBorderColor,
//							align,valign);
//	}
//	
//	/**
//	 * 変換
//	 * @param enums
//	 * @return
//	 */
//	private String arrayToString(Enum<?> ...enums) {
//		StringBuilder builder = new StringBuilder();
//		for (Enum<?> item : enums) {
//			if (item instanceof Color) {
//				builder.append( 
//						PoiUtil.nvl( (Color)item , Color.BLACK ).value()
//						);
//			}
//			else if (item instanceof FillPattern) {
//				builder.append(
//						PoiUtil.nvl( (FillPattern)item , FillPattern.NO_FILL ).value()
//						);
//			}
//			else if (item instanceof Border) {
//				builder.append(
//						PoiUtil.nvl( (Border)item , Border.BORDER_NONE ).value()
//						);
//			}
//			else if(item instanceof Align) {
//				builder.append(
//						PoiUtil.nvl( (Align)item , Align.ALIGN_GENERAL).value()
//						);
//			}
//			else if(item instanceof VAlign) {
//				builder.append(
//						PoiUtil.nvl( (VAlign)item ,  VAlign.VERTICAL_CENTER ).value()
//						);
//			}
//			builder.append(COLON);
//		}
//		
//		return builder.toString();
//	}
	
}