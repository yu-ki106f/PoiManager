package org.poco.framework.poi.factory;

import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiRange;
import org.poco.framework.poi.managers.impl.StyleManager;

public class StyleManagerFactory {
	/**
	 * 総称設定用 インスタンス取得
	 * @param value
	 * @return
	 */
	public static IStyleManager<IPoiCell> create(IPoiCell value) {
		StyleManager<IPoiCell> instance = new StyleManager<IPoiCell>(value);
		instance.parent = value;
		return instance;
	}
	
	/**
	 * 総称設定用 インスタンス取得
	 * @param value
	 * @return
	 */
	public static IStyleManager<IPoiCell> create(IPoiCell value, PoiStyleDto dto) {
		StyleManager<IPoiCell> instance = new StyleManager<IPoiCell>(value,dto);
		instance.parent = value;
		return instance;
	}
	
	/**
	 * 総称設定用 インスタンス取得
	 * @param value
	 * @return
	 */
	public static IStyleManager<IPoiRange> create(IPoiRange value) {
		StyleManager<IPoiRange> instance = new StyleManager<IPoiRange>(value);
		instance.parent = value;
		return instance;
	}
	
	/**
	 * 総称設定用 インスタンス取得
	 * @param value
	 * @return
	 */
	public static IStyleManager<IPoiRange> create(IPoiRange value, PoiStyleDto dto) {
		StyleManager<IPoiRange> instance = new StyleManager<IPoiRange>(value, dto);
		instance.parent = value;
		return instance;
	}
}
