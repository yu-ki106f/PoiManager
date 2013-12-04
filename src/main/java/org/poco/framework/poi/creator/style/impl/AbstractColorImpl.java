package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants;
import org.poco.framework.poi.constants.PoiConstants.Color;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IColor;

public abstract class AbstractColorImpl<T> implements IColor<T> {
	/**
	 * 更新処理を実装してください。
	 * @param color
	 */
	abstract protected void update(Color color);


	public IStyleManager<T> clear() {
		update(null);
		return getStyle();
	}
	

	public IStyleManager<T> AQUA() {
		update(PoiConstants.Color.AQUA);
		return getStyle();
	}


	public IStyleManager<T> BLACK() {
		update(PoiConstants.Color.BLACK);
		return getStyle();
	}


	public IStyleManager<T> BLUE() {
		update(PoiConstants.Color.BLUE);
		return getStyle();
	}


	public IStyleManager<T> BLUE_GREY() {
		update(PoiConstants.Color.BLUE_GREY);
		return getStyle();
	}


	public IStyleManager<T> BRIGHT_GREEN() {
		update(PoiConstants.Color.BRIGHT_GREEN);
		return getStyle();
	}


	public IStyleManager<T> BROWN() {
		update(PoiConstants.Color.BROWN);
		return getStyle();
	}


	public IStyleManager<T> CORAL() {
		update(PoiConstants.Color.CORAL);
		return getStyle();
	}


	public IStyleManager<T> CORNFLOWER_BLUE() {
		update(PoiConstants.Color.CORNFLOWER_BLUE);
		return getStyle();
	}


	public IStyleManager<T> DARK_BLUE() {
		update(PoiConstants.Color.DARK_BLUE);
		return getStyle();
	}


	public IStyleManager<T> DARK_GREEN() {
		update(PoiConstants.Color.DARK_GREEN);
		return getStyle();
	}


	public IStyleManager<T> DARK_RED() {
		update(PoiConstants.Color.DARK_RED);
		return getStyle();
	}


	public IStyleManager<T> DARK_TEAL() {
		update(PoiConstants.Color.DARK_TEAL);
		return getStyle();
	}


	public IStyleManager<T> DARK_YELLOW() {
		update(PoiConstants.Color.DARK_YELLOW);
		return getStyle();
	}


	public IStyleManager<T> GOLD() {
		update(PoiConstants.Color.GOLD);
		return getStyle();
	}


	public IStyleManager<T> GREEN() {
		update(PoiConstants.Color.GREEN);
		return getStyle();
	}


	public IStyleManager<T> GREY_25_PERCENT() {
		update(PoiConstants.Color.GREY_25_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> GREY_40_PERCENT() {
		update(PoiConstants.Color.GREY_40_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> GREY_50_PERCENT() {
		update(PoiConstants.Color.GREY_50_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> GREY_80_PERCENT() {
		update(PoiConstants.Color.GREY_80_PERCENT);
		return getStyle();
	}


	public IStyleManager<T> INDIGO() {
		update(PoiConstants.Color.INDIGO);
		return getStyle();
	}


	public IStyleManager<T> LAVENDER() {
		update(PoiConstants.Color.LAVENDER);
		return getStyle();
	}


	public IStyleManager<T> LEMON_CHIFFON() {
		update(PoiConstants.Color.LEMON_CHIFFON);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_BLUE() {
		update(PoiConstants.Color.LIGHT_BLUE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_CORNFLOWER_BLUE() {
		update(PoiConstants.Color.LIGHT_CORNFLOWER_BLUE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_GREEN() {
		update(PoiConstants.Color.LIGHT_GREEN);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_ORANGE() {
		update(PoiConstants.Color.LIGHT_ORANGE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_TURQUOISE() {
		update(PoiConstants.Color.LIGHT_TURQUOISE);
		return getStyle();
	}


	public IStyleManager<T> LIGHT_YELLOW() {
		update(PoiConstants.Color.LIGHT_YELLOW);
		return getStyle();
	}


	public IStyleManager<T> LIME() {
		update(PoiConstants.Color.LIME);
		return getStyle();
	}


	public IStyleManager<T> MAROON() {
		update(PoiConstants.Color.MAROON);
		return getStyle();
	}


	public IStyleManager<T> OLIVE_GREEN() {
		update(PoiConstants.Color.OLIVE_GREEN);
		return getStyle();
	}


	public IStyleManager<T> ORANGE() {
		update(PoiConstants.Color.ORANGE);
		return getStyle();
	}


	public IStyleManager<T> ORCHID() {
		update(PoiConstants.Color.ORCHID);
		return getStyle();
	}


	public IStyleManager<T> PALE_BLUE() {
		update(PoiConstants.Color.PALE_BLUE);
		return getStyle();
	}


	public IStyleManager<T> PINK() {
		update(PoiConstants.Color.PINK);
		return getStyle();
	}


	public IStyleManager<T> PLUM() {
		update(PoiConstants.Color.PLUM);
		return getStyle();
	}


	public IStyleManager<T> RED() {
		update(PoiConstants.Color.RED);
		return getStyle();
	}


	public IStyleManager<T> ROSE() {
		update(PoiConstants.Color.ROSE);
		return getStyle();
	}


	public IStyleManager<T> ROYAL_BLUE() {
		update(PoiConstants.Color.ROYAL_BLUE);
		return getStyle();
	}


	public IStyleManager<T> SEA_GREEN() {
		update(PoiConstants.Color.SEA_GREEN);
		return getStyle();
	}


	public IStyleManager<T> SKY_BLUE() {
		update(PoiConstants.Color.SKY_BLUE);
		return getStyle();
	}


	public IStyleManager<T> TAN() {
		update(PoiConstants.Color.TAN);
		return getStyle();
	}


	public IStyleManager<T> TEAL() {
		update(PoiConstants.Color.TEAL);
		return getStyle();
	}


	public IStyleManager<T> TURQUOISE() {
		update(PoiConstants.Color.TURQUOISE);
		return getStyle();
	}


	public IStyleManager<T> VIOLET() {
		update(PoiConstants.Color.VIOLET);
		return getStyle();
	}


	public IStyleManager<T> WHITE() {
		update(PoiConstants.Color.WHITE);
		return getStyle();
	}


	public IStyleManager<T> YELLOW() {
		update(PoiConstants.Color.YELLOW);
		return getStyle();
	}
}
