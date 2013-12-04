/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants;
import org.poco.framework.poi.constants.PoiConstants.Align;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IAlignStyle;

/**
 * @author yu-ki106f
 *
 */
public abstract class AbstractAlignImpl<T> implements IAlignStyle<T> {

	abstract protected void update(Align align);
	

	public IStyleManager<T> ALIGN_CENTER() {
		update(PoiConstants.Align.ALIGN_CENTER);
		return getStyle();
	}


	public IStyleManager<T> ALIGN_CENTER_SELECTION() {
		update(PoiConstants.Align.ALIGN_CENTER_SELECTION);
		return getStyle();
	}


	public IStyleManager<T> ALIGN_FILL() {
		update(PoiConstants.Align.ALIGN_FILL);
		return getStyle();
	}


	public IStyleManager<T> ALIGN_GENERAL() {
		update(PoiConstants.Align.ALIGN_GENERAL);
		return getStyle();
	}


	public IStyleManager<T> ALIGN_JUSTIFY() {
		update(PoiConstants.Align.ALIGN_JUSTIFY);
		return getStyle();
	}


	public IStyleManager<T> ALIGN_LEFT() {
		update(PoiConstants.Align.ALIGN_LEFT);
		return getStyle();
	}


	public IStyleManager<T> ALIGN_RIGHT() {
		update(PoiConstants.Align.ALIGN_RIGHT);
		return getStyle();
	}

}
