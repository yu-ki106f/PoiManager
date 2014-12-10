package org.poco.framework.poi.dto;

import java.lang.reflect.Field;

import org.poco.framework.poi.constants.PoiConstants.BoldWeight;
import org.poco.framework.poi.constants.PoiConstants.FLColor;
import org.poco.framework.poi.constants.PoiConstants.FontOffset;
import org.poco.framework.poi.constants.PoiConstants.FontUnderline;
import org.poco.framework.poi.utils.StyleUtil;

import org.apache.poi.ss.usermodel.Font;

public class PoiFontDto extends AbstractDto {
	
	/**
	 * フォント名
	 */
	public String fontName;
	
	/**
	 * イタリック
	 */
	public Boolean italic;
	
	/**
	 * フォント幅
	 */
	public Short fontHeight;
	
	/**
	 * ポイント
	 */
	public Short fontHeightInPoints;
	
	/**
	 * 太さ
	 */
	public BoldWeight boldweight;
	
	/**
	 * 色
	 */
	public FLColor color;
	
	/**
	 * 取り消し線
	 */
	public Boolean strikeout;
	
	/**
	 * アンダーライン
	 */
	public FontUnderline underline;
	
	/**
	 * 上付き・下付き
	 */
	public FontOffset typeOffset;

	private Font font = null;
	
	public Font getOrgFont() {
		return font;
	}
	
	/**
	 * コンストラクタ
	 * @param font
	 */
	public PoiFontDto() {}
	
	/**
	 * コンストラクタ
	 * @param font
	 */
	public PoiFontDto(Font font) {
		if (font == null) return;
		
		boldweight = StyleUtil.getBoldWeight(font.getBoldweight());
		color = StyleUtil.getColor(font.getColor());
		fontHeight = font.getFontHeight();
		fontHeightInPoints = font.getFontHeightInPoints();
		fontName = font.getFontName();
		italic = font.getItalic();
		strikeout = font.getStrikeout();
		typeOffset = StyleUtil.getFontOffset(font.getTypeOffset());
		underline = StyleUtil.getFontUnderline(font.getUnderline());
		
		this.font = font;
	}

//	/**
//	 * 比較
//	 * @param dto
//	 * @return
//	 */
//	public boolean equals(PoiFontDto dto) {
//		for (Field f : this.getClass().getFields()) {
//			try {
//				if (!f.get(this).equals(f.get(dto)) ) {
//					return false;
//				}
//			} catch (Exception e) {
//				return false;
//			}
//		}
//		return true;
//	}

	/**
	 * 更新
	 * @param font
	 */
	public void update(Font font) {
		if (boldweight != null) {
			font.setBoldweight(boldweight.value());
		}
		if (color != null) {
			font.setColor(color.value());
		}
		if (fontHeight != null) {
			font.setFontHeight(fontHeight);
		}
		if (fontHeightInPoints != null) {
			font.setFontHeightInPoints(fontHeightInPoints);
		}
		if (fontName != null) {
			font.setFontName(fontName);
			font.setCharSet(Font.DEFAULT_CHARSET);
		}
		if (italic != null) {
			font.setItalic(italic);
		}
		if (strikeout != null) {
			font.setStrikeout(strikeout);
		}
		if (typeOffset != null) {
			font.setTypeOffset(typeOffset.value());
		}
		if (underline != null) {
			font.setUnderline(underline.value());
		}
	}
	
	public PoiFontDto merge(PoiFontDto dto) {
		//同じならマージしない
		if (this.font.equals(dto.getOrgFont())) return this;
		
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

}
