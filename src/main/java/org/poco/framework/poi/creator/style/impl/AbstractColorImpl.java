package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants;
import org.poco.framework.poi.constants.PoiConstants.FLColor;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IColor;

public abstract class AbstractColorImpl<T> implements IColor<T> {
	/**
	 * 更新処理を実装してください。
	 * @param color
	 */
	abstract protected void update(FLColor color);


	public IStyleManager<T> clear() {
		update(null);
		return getStyle();
	}
	

	public IStyleManager<T> AQUA() {
		update(PoiConstants.FLColor.AQUA);
		return getStyle();
	}


	public IStyleManager<T> BLACK() {
		update(PoiConstants.FLColor.BLACK);
		return getStyle();
	}


	public IStyleManager<T> BLUE() {
		update(PoiConstants.FLColor.BLUE);
		return getStyle();
	}


	public IStyleManager<T> BLUE_GREY() {
		update(PoiConstants.FLColor.BLUE_GREY);
		return getStyle();
	}


	public IStyleManager<T> BRIGHT_GREEN() {
		update(PoiConstants.FLColor.BRIGHT_GREEN);
		return getStyle();
	}


	public IStyleManager<T> BROWN() {
		update(PoiConstants.FLColor.BROWN);
		return getStyle();
	}


	public IStyleManager<T> CORAL() {
		update(PoiConstants.FLColor.CORAL);
		return getStyle();
	}


	public IStyleManager<T> CORNFLOWER_BLUE() {
		update(PoiConstants.FLColor.CORNFLOWER_BLUE);
		return getStyle();
	}


	public IStyleManager<T> DARK_BLUE() {
		update(PoiConstants.FLColor.DARK_BLUE);
		return getStyle();
	}


	public IStyleManager<T> DARK_GREEN() {
		update(PoiConstants.FLColor.DARK_GREEN);
		return getStyle();
	}


	public IStyleManager<T> DARK_RED() {
		update(PoiConstants.FLColor.DARK_RED);
		return getStyle();
	}


	public IStyleManager<T> DARK_TEAL() {
		update(PoiConstants.FLColor.DARK_TEAL);
		return getStyle();
	}


	public IStyleManager<T> DARK_YELLOW() {
		update(PoiConstants.FLColor.DARK_YELLOW);
		return getStyle();
	}


	public IStyleManager<T> GOLD() {
		update(PoiConstants.FLColor.GOLD);
		return getStyle();
	}


	public IStyleManager<T> GREEN() {
		update(PoiConstants.FLColor.GREEN);
		return getStyle();
	}


	public IStyleManager<T> GREY_25_PERCENT() {
		update(PoiConstants.FLColor.GREY_25_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> GREY_40_PERCENT() {
		update(PoiConstants.FLColor.GREY_40_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> GREY_50_PERCENT() {
		update(PoiConstants.FLColor.GREY_50_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> GREY_80_PERCENT() {
		update(PoiConstants.FLColor.GREY_80_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> INDIGO() {
		update(PoiConstants.FLColor.INDIGO);
		return getStyle();
	}


	public IStyleManager<T> LAVENDER() {
		update(PoiConstants.FLColor.LAVENDER);
		return getStyle();
	}


	public IStyleManager<T> LEMON_CHIFFON() {
		update(PoiConstants.FLColor.LEMON_CHIFFON);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_BLUE() {
		update(PoiConstants.FLColor.LIGHT_BLUE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_CORNFLOWER_BLUE() {
		update(PoiConstants.FLColor.LIGHT_CORNFLOWER_BLUE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_GREEN() {
		update(PoiConstants.FLColor.LIGHT_GREEN);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_ORANGE() {
		update(PoiConstants.FLColor.LIGHT_ORANGE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_TURQUOISE() {
		update(PoiConstants.FLColor.LIGHT_TURQUOISE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_YELLOW() {
		update(PoiConstants.FLColor.LIGHT_YELLOW);
		return getStyle();
	}


	public IStyleManager<T> LIME() {
		update(PoiConstants.FLColor.LIME);
		return getStyle();
	}


	public IStyleManager<T> MAROON() {
		update(PoiConstants.FLColor.MAROON);
		return getStyle();
	}


	public IStyleManager<T> OLIVE_GREEN() {
		update(PoiConstants.FLColor.OLIVE_GREEN);
		return getStyle();
	}


	public IStyleManager<T> ORANGE() {
		update(PoiConstants.FLColor.ORANGE);
		return getStyle();
	}


	public IStyleManager<T> ORCHID() {
		update(PoiConstants.FLColor.ORCHID);
		return getStyle();
	}


	public IStyleManager<T> PALE_BLUE() {
		update(PoiConstants.FLColor.PALE_BLUE);
		return getStyle();
	}


	public IStyleManager<T> PINK() {
		update(PoiConstants.FLColor.PINK);
		return getStyle();
	}


	public IStyleManager<T> PLUM() {
		update(PoiConstants.FLColor.PLUM);
		return getStyle();
	}


	public IStyleManager<T> RED() {
		update(PoiConstants.FLColor.RED);
		return getStyle();
	}


	public IStyleManager<T> ROSE() {
		update(PoiConstants.FLColor.ROSE);
		return getStyle();
	}


	public IStyleManager<T> ROYAL_BLUE() {
		update(PoiConstants.FLColor.ROYAL_BLUE);
		return getStyle();
	}


	public IStyleManager<T> SEA_GREEN() {
		update(PoiConstants.FLColor.SEA_GREEN);
		return getStyle();
	}


	public IStyleManager<T> SKY_BLUE() {
		update(PoiConstants.FLColor.SKY_BLUE);
		return getStyle();
	}


	public IStyleManager<T> TAN() {
		update(PoiConstants.FLColor.TAN);
		return getStyle();
	}


	public IStyleManager<T> TEAL() {
		update(PoiConstants.FLColor.TEAL);
		return getStyle();
	}


	public IStyleManager<T> TURQUOISE() {
		update(PoiConstants.FLColor.TURQUOISE);
		return getStyle();
	}


	public IStyleManager<T> VIOLET() {
		update(PoiConstants.FLColor.VIOLET);
		return getStyle();
	}


	public IStyleManager<T> WHITE() {
		update(PoiConstants.FLColor.WHITE);
		return getStyle();
	}


	public IStyleManager<T> YELLOW() {
		update(PoiConstants.FLColor.YELLOW);
		return getStyle();
	}
}
