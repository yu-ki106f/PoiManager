/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants;
import org.poco.framework.poi.constants.PoiConstants.FillPattern;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IFillPattern;

/**
 * @author yu-ki106f
 *
 */
public abstract class AbstractFillPatternImpl<T> implements IFillPattern<T> {

	abstract protected void update(FillPattern fill);
	

	public IStyleManager<T> ALT_BARS() {
		update(PoiConstants.FillPattern.ALT_BARS);
		return getStyle();
	}


	public IStyleManager<T> BIG_SPOTS() {
		update(PoiConstants.FillPattern.BIG_SPOTS);
		return getStyle();
	}


	public IStyleManager<T> BRICKS() {
		update(PoiConstants.FillPattern.BRICKS);
		return getStyle();
	}


	public IStyleManager<T> DIAMONDS() {
		update(PoiConstants.FillPattern.DIAMONDS);
		return getStyle();
	}


	public IStyleManager<T> FINE_DOTS() {
		update(PoiConstants.FillPattern.FINE_DOTS);
		return getStyle();
	}


	public IStyleManager<T> NO_FILL() {
		update(PoiConstants.FillPattern.NO_FILL);
		return getStyle();
	}


	public IStyleManager<T> SOLID_FOREGROUND() {
		update(PoiConstants.FillPattern.SOLID_FOREGROUND);
		return getStyle();
	}


	public IStyleManager<T> SPARSE_DOTS() {
		update(PoiConstants.FillPattern.SPARSE_DOTS);
		return getStyle();
	}


	public IStyleManager<T> SQUARES() {
		update(PoiConstants.FillPattern.SQUARES);
		return getStyle();
	}


	public IStyleManager<T> THICK_BACKWARD_DIAG() {
		update(PoiConstants.FillPattern.THICK_BACKWARD_DIAG);
		return getStyle();
	}


	public IStyleManager<T> THICK_FORWARD_DIAG() {
		update(PoiConstants.FillPattern.THICK_FORWARD_DIAG);
		return getStyle();
	}


	public IStyleManager<T> THICK_HORZ_BANDS() {
		update(PoiConstants.FillPattern.THICK_HORZ_BANDS);
		return getStyle();
	}


	public IStyleManager<T> THICK_VERT_BANDS() {
		update(PoiConstants.FillPattern.THICK_VERT_BANDS);
		return getStyle();
	}


	public IStyleManager<T> THIN_BACKWARD_DIAG() {
		update(PoiConstants.FillPattern.THIN_BACKWARD_DIAG);
		return getStyle();
	}


	public IStyleManager<T> THIN_FORWARD_DIAG() {
		update(PoiConstants.FillPattern.THIN_FORWARD_DIAG);
		return getStyle();
	}


	public IStyleManager<T> THIN_HORZ_BANDS() {
		update(PoiConstants.FillPattern.THIN_HORZ_BANDS);
		return getStyle();
	}


	public IStyleManager<T> THIN_VERT_BANDS() {
		update(PoiConstants.FillPattern.THIN_VERT_BANDS);
		return getStyle();
	}

}
