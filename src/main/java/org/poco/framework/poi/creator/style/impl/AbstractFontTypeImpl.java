/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.BoldWeight;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IFontType;

/**
 * @author yu-ki106f
 *
 */
public abstract class AbstractFontTypeImpl<T> implements IFontType<T> {

	abstract protected void update(BoldWeight boldWeight );
	

	public IStyleManager<T> BOLD() {
		update(BoldWeight.BOLDWEIGHT_BOLD);
		return getStyle();
	}


	public IStyleManager<T> NORMAL() {
		update(BoldWeight.BOLDWEIGHT_NORMAL);
		return getStyle();
	}
}
