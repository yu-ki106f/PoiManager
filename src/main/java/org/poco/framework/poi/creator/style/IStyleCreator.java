package org.poco.framework.poi.creator.style;

import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;

public interface IStyleCreator<T> {
	/**
	 * スタイル管理インスタンスを取得します。
	 * @return
	 */
	IStyleManager<T> getStyle();

	/**
	 * StyleDtoを取得します。
	 * @return
	 */
	PoiStyleDto getStyleDto();
}
