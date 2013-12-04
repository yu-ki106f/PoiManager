/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants;
import org.poco.framework.poi.constants.PoiConstants.VAlign;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IVAlignStyle;

/**
 * @author yu-ki106f
 *
 */
public abstract class AbstractVAlignImpl<T> implements IVAlignStyle<T> {

	abstract protected void update(VAlign valign);


	public IStyleManager<T> VERTICAL_BOTTOM() {
		update(PoiConstants.VAlign.VERTICAL_BOTTOM);
		return getStyle();
	}


	public IStyleManager<T> VERTICAL_CENTER() {
		update(PoiConstants.VAlign.VERTICAL_CENTER);
		return getStyle();
	}


	public IStyleManager<T> VERTICAL_JUSTIFY() {
		update(PoiConstants.VAlign.VERTICAL_JUSTIFY);
		return getStyle();
	}


	public IStyleManager<T> VERTICAL_TOP() {
		update(PoiConstants.VAlign.VERTICAL_TOP);
		return getStyle();
	}

}
