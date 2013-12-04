package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants;
import org.poco.framework.poi.constants.PoiConstants.Border;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IBorderStyle;

/**
 * @author yu-ki106f
 *
 */
public abstract class AbstractBorderImpl<T> implements IBorderStyle<T> {

	abstract protected void update(Border border);
	

	public IStyleManager<T> clear() {
		update(null);
		return getStyle();
	}
	

	public IStyleManager<T> BORDER_DASHED() {
		update(PoiConstants.Border.BORDER_DASHED);
		return getStyle();
	}


	public IStyleManager<T> BORDER_DASH_DOT() {
		update(PoiConstants.Border.BORDER_DASH_DOT);
		return getStyle();
	}


	public IStyleManager<T> BORDER_DASH_DOT_DOT() {
		update(PoiConstants.Border.BORDER_DASH_DOT_DOT);
		return getStyle();
	}


	public IStyleManager<T> BORDER_DOTTED() {
		update(PoiConstants.Border.BORDER_DOTTED);
		return getStyle();
	}


	public IStyleManager<T> BORDER_DOUBLE() {
		update(PoiConstants.Border.BORDER_DOUBLE);
		return getStyle();
	}


	public IStyleManager<T> BORDER_HAIR() {
		update(PoiConstants.Border.BORDER_HAIR);
		return getStyle();
	}


	public IStyleManager<T> BORDER_MEDIUM() {
		update(PoiConstants.Border.BORDER_MEDIUM);
		return getStyle();
	}


	public IStyleManager<T> BORDER_MEDIUM_DASHED() {
		update(PoiConstants.Border.BORDER_MEDIUM_DASHED);
		return getStyle();
	}


	public IStyleManager<T> BORDER_MEDIUM_DASH_DOT() {
		update(PoiConstants.Border.BORDER_MEDIUM_DASH_DOT);
		return getStyle();
	}


	public IStyleManager<T> BORDER_MEDIUM_DASH_DOT_DOT() {
		update(PoiConstants.Border.BORDER_MEDIUM_DASH_DOT_DOT);
		return getStyle();
	}


	public IStyleManager<T> BORDER_NONE() {
		update(PoiConstants.Border.BORDER_NONE);
		return getStyle();
	}


	public IStyleManager<T> BORDER_SLANTED_DASH_DOT() {
		update(PoiConstants.Border.BORDER_SLANTED_DASH_DOT);
		return getStyle();
	}


	public IStyleManager<T> BORDER_THICK() {
		update(PoiConstants.Border.BORDER_THICK);
		return getStyle();
	}


	public IStyleManager<T> BORDER_THIN() {
		update(PoiConstants.Border.BORDER_THIN);
		return getStyle();
	}

}
