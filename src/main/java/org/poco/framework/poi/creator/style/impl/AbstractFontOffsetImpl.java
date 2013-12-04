package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.FontOffset;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IFontOffset;

public abstract class AbstractFontOffsetImpl<T> implements IFontOffset<T> {

	abstract protected void update(FontOffset offset);
	

	public IStyleManager<T> SS_NONE() {
		update(FontOffset.SS_NONE);
		return getStyle();
	}


	public IStyleManager<T> SS_SUB() {
		update(FontOffset.SS_SUB);
		return getStyle();
	}


	public IStyleManager<T> SS_SUPER() {
		update(FontOffset.SS_SUPER);
		return getStyle();
	}

}
