package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.creator.style.IStyleCreator;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;

public abstract class StyleCreatorImpl<T> implements IStyleCreator<T> {

	private PoiStyleDto styleDto;

	private IStyleManager<T> style;
	
	/**
	 * コンストラクタ
	 * @param dto
	 */
	public StyleCreatorImpl(IStyleManager<T> style, PoiStyleDto dto) {
		this.style = style;
		this.styleDto = dto;
	}

	/**
	 * スタイル管理インスタンスを取得します。
	 * @return
	 */
	public IStyleManager<T> getStyle() {
		return style;
	}

	/**
	 * StyleDtoを取得します。
	 * @return
	 */
	public PoiStyleDto getStyleDto() {
		return styleDto;
	}

}
