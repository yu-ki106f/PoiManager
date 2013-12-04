package org.poco.framework.poi.factory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.poco.framework.poi.dto.PoiFontDto;
import org.poco.framework.poi.dto.PoiStyleDto;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * キャッシュするのでメモリ開放が必要
 * @author yu-ki106f
 *
 */
public class StyleFactory {
	
	private static Map<HSSFWorkbook,StyleFactory> factoryCashe = new HashMap<HSSFWorkbook, StyleFactory>();
	
	private List<PoiStyleDto> cashe = new ArrayList<PoiStyleDto>();
	private List<PoiFontDto> fontCashe = new ArrayList<PoiFontDto>();
	
	private HSSFWorkbook book = null;

	public static class IsFont {
		private Boolean value = false;
		public void setValue(Boolean value) {
			this.value = value;
		}
		public Boolean getValue() {
			return value;
		}
	}
	
	/**
	 * 開放
	 */
	public static void destory() {
		factoryCashe.clear();
	}
	
	/**
	 * インスタンスを取得
	 * @param book
	 * @return
	 */
	public static StyleFactory getInstance(HSSFWorkbook book) {
		StyleFactory factory = factoryCashe.get(book);
		if (factory == null) {
			factory = new StyleFactory(book);
			factoryCashe.put(book,factory);
		}
		return factory;
	}
	
	/**
	 * コンストラクタ
	 * @param book
	 */
	private StyleFactory(HSSFWorkbook book) {
		this.book = book;

		PoiFontDto dto;
		for (short idx = 0; idx < book.getNumberOfFonts(); idx++ ) {
			dto = new PoiFontDto(book.getFontAt(idx));
			if (searchFont(dto)==null) {
				fontCashe.add(0,dto);
			}
		}
		
		PoiStyleDto dto2;
		IsFont isFont = new IsFont();
		for (short idx = 0; idx < book.getNumCellStyles(); idx++ ) {
			dto2 = new PoiStyleDto(book.getCellStyleAt(idx), book.getFontAt(book.getCellStyleAt(idx).getFontIndex()) );
			if (searchStyle(dto2,isFont) == null) {
				if (isFont.getValue() == false) {
					fontCashe.add(0,dto2.getFont());
				}
				cashe.add(0,dto2);
			}
		}
	}
	
	/**
	 * fontを検索
	 * @param param
	 * @return
	 */
	private HSSFFont searchFont(PoiFontDto param) {

		for (PoiFontDto fontDto: fontCashe) {
			if (fontDto.equals(param)) {
				return fontDto.getOrgFont();
			}
		}
		return null;
	}
	
	/**
	 * styleを検索
	 * @param param
	 * @return
	 */
	private HSSFCellStyle searchStyle(PoiStyleDto param, IsFont isFont) {
		HSSFFont font = searchFont(param.getFont());
		if (font == null) {
			isFont.setValue(false);
			return null;
		}
		//FONTはある
		isFont.setValue(true);

		for (PoiStyleDto dto : cashe) {
			if (dto.equals(param)) {
				return dto.getStyle();
			}
		}
		return null;
	}
	
	public void addStyleCashe(PoiStyleDto dto) {
		//なければ追加
		if (!cashe.contains(dto)) {
			cashe.add(0,dto);
		}
	}
	
	public void addFontCashe(PoiFontDto dto) {
		//なければ追加
		if (!fontCashe.contains(dto)) {
			fontCashe.add(0,dto);
		}
	}
	
	/**
	 * オリジナルからのマージ処理
	 * @param src 設定内容
	 * @param orgCellStyle 元のセルスタイル
	 * @return StyleDto
	 */
	private PoiStyleDto mergeStyle(PoiStyleDto src, HSSFCellStyle orgCellStyle)
	{
		PoiFontDto font = new PoiFontDto(orgCellStyle.getFont(this.book));

		//マージ
		font.merge(src.getFont());

		PoiStyleDto style = new PoiStyleDto(orgCellStyle, null);
		//マージ
		style.merge(src);
		
		//フォントを設定
		style.setFont(font);
		
		//結果
		return style;
	}
	
	/**
	 * スタイルを取得
	 * @param dto
	 * @param orgCellStyle
	 * @return
	 */
	public /*synchronized*/ HSSFCellStyle getStyle(PoiStyleDto dto, HSSFCellStyle orgCellStyle) {
		PoiStyleDto merge = mergeStyle(dto, orgCellStyle);
		IsFont isFont = new IsFont();
		HSSFCellStyle style = searchStyle(merge, isFont);
		//存在しない場合は、作成する
		if (style == null) {
			HSSFFont font = merge.getFont().getOrgFont();
			//ない場合
			if (isFont.getValue() == false) {
				//作成
				font = createFont(merge.getFont());
			}
			//作成
			style = createStyle(font,merge);
		}
		return style;
	}
	
	/**
	 * フォント作成
	 * @param book
	 * @return
	 */
	public HSSFFont createFont(PoiFontDto updateDto) {
		HSSFFont result = book.createFont();
		updateDto.update(result);
		PoiFontDto dto = new PoiFontDto(result);
		//キャッシュ追加
		if (searchFont(dto) == null) {
			addFontCashe(dto);
		}
		return result;
	}

	/**
	 * スタイル作成
	 * @param book
	 * @param font
	 * @return
	 */
	public HSSFCellStyle createStyle(HSSFFont font,PoiStyleDto updateDto) {
		
		HSSFCellStyle result = book.createCellStyle();
		updateDto.update(result,book);
		result.setFont(font);
		
		PoiStyleDto dto = new PoiStyleDto(result,font);
		IsFont isFont = new IsFont();
		//キャッシュ追加
		if (searchStyle(dto,isFont) == null) {
			if (isFont.getValue() == false) {
				createFont(dto.getFont());
			}
			addStyleCashe(dto);
		}
		return result;
	}

}
