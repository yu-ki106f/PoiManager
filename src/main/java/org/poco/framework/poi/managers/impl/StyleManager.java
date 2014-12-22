/**
 * 
 */
package org.poco.framework.poi.managers.impl;

import org.poco.framework.poi.creator.style.impl.AlignStyleCreator;
import org.poco.framework.poi.creator.style.impl.BorderStyleCreator;
import org.poco.framework.poi.creator.style.impl.FillStyleCreator;
import org.poco.framework.poi.creator.style.impl.FontStyleCreator;
import org.poco.framework.poi.creator.style.impl.VAlignStyleCreator;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.factory.StyleFactory;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IPoiManager.IPoiBook;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiRange;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author yu-ki
 *
 */
public class StyleManager<T> implements IStyleManager<T> {
	
	private static final String DEFAULT_DATA_FORMAT = "General";
	
	private IPoiCell[] cells;
	private PoiStyleDto styleDto;
	
	public  T parent = null;
	
	/**
	 * コンストラクタ
	 * @param value
	 */
	public StyleManager(IPoiCell value) {
		this(value, 
			new PoiStyleDto()
		);

	}

	/**
	 * コンストラクタ
	 * @param value
	 * @param dto
	 */
	public StyleManager(IPoiCell value, PoiStyleDto dto) {
		this.cells = new IPoiCell[1];
		this.cells[0] = value;
		this.styleDto = dto;
	}

	/**
	 * コンストラクタ
	 * @param value
	 */
	public StyleManager(IPoiRange value) {
		this(value, 
				new PoiStyleDto()
				);
	}
	
	/**
	 * コンストラクタ
	 * @param value
	 * @param dto
	 */
	@SuppressWarnings("unchecked")
	public StyleManager(IPoiRange value, PoiStyleDto dto) {
		this.cells = value.getCells();
		this.styleDto = dto;
		
		this.parent = (T)value;
	}
	
	/**
	 * ワークブック
	 * @return
	 */
	public Workbook getWorkBook() {
		return cells[0].getOrgCell().getSheet().getWorkbook();
	}
	
	public IPoiBook getPoiBook() {
		return cells[0].getBook();
	}

	/**
	 * Cell 先頭1件
	 * @return
	 */
	public Cell getCell() {
		return getCellAt(0);
	}
	
	public Cell getCellAt(int index) {
		return cells[index].getOrgCell();
	}
	/**
	 * Cell 全件
	 * @return
	 */
	public Cell[] getCells() {
		int length = cells.length;
		Cell[] result = new Cell[length];
		for (int i=0; i<length;i++) {
			result[i] = cells[i].getOrgCell();
		}
		return result;
	}
	
	public IAlignStyle<T> align() {
		return new AlignStyleCreator<T>(this,styleDto).createAlignStyle();
	}

	public IBorderType<T> border() {
		return new BorderStyleCreator<T>(this,styleDto).createBorderType();
	}

	public IFontType<T> font() {
		return new FontStyleCreator<T>(this,styleDto).createFontType();
	}

	public IVAlignStyle<T> valign() {
		return new VAlignStyleCreator<T>(this,styleDto).createVAlignStyle();
	}

	public IFillType<T> fill() {
		return new FillStyleCreator<T>(this,styleDto).createFillType();
	}

	public StyleManager<T> setDataFormat(String value) {
		if (value == null || value.length() == 0) {
			styleDto.dataFormat = DEFAULT_DATA_FORMAT;
		}
		else {
			styleDto.dataFormat = value;
		}
		return this;
	}
	
	public T update() {
		CellStyle style = null;
		//何もしていないCellは、生成しないようにIPoiCellループに変更
		for (IPoiCell cell : cells) {
			//スタイル取得
			style = StyleFactory.getInstance(getPoiBook()).getStyle(styleDto, cell.getOrgCell().getCellStyle());
			//反映
			cell.getOrgCell().setCellStyle(style);

		}
		return parent;
	}
	public T forceUpdate(CellStyle style) {
		//何もしていないCellは、生成しないようにIPoiCellループに変更
		for (IPoiCell cell : cells) {
			//反映
			cell.getOrgCell().setCellStyle(style);
		}
		return parent;
	}
}
