package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.FontUnderline;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IFontUnderline;

public abstract class AbstractFontUnderlineImpl<T> implements IFontUnderline<T> {

	abstract protected void update(FontUnderline underline);
	

	public IStyleManager<T> U_DOUBLE() {
		update(FontUnderline.U_DOUBLE);
		return getStyle();
	}


	public IStyleManager<T> U_DOUBLE_ACCOUNTING() {
		update(FontUnderline.U_DOUBLE_ACCOUNTING);
		return getStyle();
	}


	public IStyleManager<T> U_NONE() {
		update(FontUnderline.U_NONE);
		return getStyle();
	}


	public IStyleManager<T> U_SINGLE() {
		update(FontUnderline.U_SINGLE);
		return getStyle();
	}


	public IStyleManager<T> U_SINGLE_ACCOUNTING() {
		update(FontUnderline.U_SINGLE_ACCOUNTING);
		return getStyle();
	}

}
