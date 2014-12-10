package org.poco.framework.poi.factory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.poco.framework.poi.dto.PoiFontDto;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IPoiManager.IPoiBook;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

/**
 * キャッシュするのでメモリ開放が必要
 * @author funahashi
 *
 */
public class StyleFactory {
	
	private static Map<IPoiBook,StyleFactory> factoryCashe = new HashMap<IPoiBook, StyleFactory>();
	
	private List<PoiStyleDto> cashe = new ArrayList<PoiStyleDto>();
	private List<PoiFontDto> fontCashe = new ArrayList<PoiFontDto>();
	
	private IPoiBook book = null;

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
	 * 指定したBookのスタイルを開放する
	 * @param book
	 */
	public static void removeStyle(IPoiBook book){
		factoryCashe.remove(book);
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
	public static StyleFactory getInstance(IPoiBook book) {
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
	private StyleFactory(IPoiBook book) {
		this.book = book;

		PoiFontDto dto;
		for (short idx = 0; idx < book.getOrgWorkBook().getNumberOfFonts(); idx++ ) {
			dto = new PoiFontDto(book.getOrgWorkBook().getFontAt(idx));
			if (searchFont(dto)==null) {
				fontCashe.add(0,dto);
			}
		}
		
		PoiStyleDto dto2;
		IsFont isFont = new IsFont();
		for (short idx = 0; idx < book.getOrgWorkBook().getNumCellStyles(); idx++ ) {
			dto2 = new PoiStyleDto(book.getOrgWorkBook().getCellStyleAt(idx), book.getOrgWorkBook().getFontAt(book.getOrgWorkBook().getCellStyleAt(idx).getFontIndex()) );
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
	private Font searchFont(PoiFontDto param) {

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
	private CellStyle searchStyle(PoiStyleDto param, IsFont isFont) {
		Font font = searchFont(param.getFont());
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
	private PoiStyleDto mergeStyle(PoiStyleDto src, CellStyle orgCellStyle)
	{
		PoiFontDto font = new PoiFontDto(this.book.getOrgWorkBook().getFontAt(orgCellStyle.getFontIndex()));

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
	public /*synchronized*/ CellStyle getStyle(PoiStyleDto dto, CellStyle orgCellStyle) {
		PoiStyleDto merge = mergeStyle(dto, orgCellStyle);
		IsFont isFont = new IsFont();
		CellStyle style = searchStyle(merge, isFont);
		//存在しない場合は、作成する
		if (style == null) {
			Font font = merge.getFont().getOrgFont();
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
	public Font createFont(PoiFontDto updateDto) {
		Font result = book.getOrgWorkBook().createFont();
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
	public CellStyle createStyle(Font font,PoiStyleDto updateDto) {
		
		CellStyle result = book.getOrgWorkBook().createCellStyle();
		updateDto.update(result,book.getOrgWorkBook());
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
