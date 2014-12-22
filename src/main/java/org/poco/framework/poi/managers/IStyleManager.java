package org.poco.framework.poi.managers;

import org.poco.framework.poi.managers.impl.StyleManager;



public interface IStyleManager<T> {
	
	IBorderType<T> border();
	
	IAlignStyle<T> align();
	
	IVAlignStyle<T> valign();
	
	IFontType<T> font();
	
	IFillType<T> fill();
	
	StyleManager<T> setDataFormat(String value);
	
	T update();
	
	public interface IColor<T> extends IStyleBase<T> {
		IStyleManager<T> clear();
		IStyleManager<T> BLACK();
		IStyleManager<T> BROWN();
		IStyleManager<T> OLIVE_GREEN();
		IStyleManager<T> DARK_GREEN();
		IStyleManager<T> DARK_TEAL();
		IStyleManager<T> DARK_BLUE();
		IStyleManager<T> INDIGO();
		IStyleManager<T> GREY_80_PERCENT();
		IStyleManager<T> ORANGE();
		IStyleManager<T> DARK_YELLOW();
		IStyleManager<T> GREEN();
		IStyleManager<T> TEAL();
		IStyleManager<T> BLUE();
		IStyleManager<T> BLUE_GREY();
		IStyleManager<T> GREY_50_PERCENT();
		IStyleManager<T> RED();
		IStyleManager<T> LIGHT_ORANGE();
		IStyleManager<T> LIME();
		IStyleManager<T> SEA_GREEN();
		IStyleManager<T> AQUA();
		IStyleManager<T> LIGHT_BLUE();
		IStyleManager<T> VIOLET();
		IStyleManager<T> GREY_40_PERCENT();
		IStyleManager<T> PINK();
		IStyleManager<T> GOLD();
		IStyleManager<T> YELLOW();
		IStyleManager<T> BRIGHT_GREEN();
		IStyleManager<T> TURQUOISE();
		IStyleManager<T> DARK_RED();
		IStyleManager<T> SKY_BLUE();
		IStyleManager<T> PLUM();
		IStyleManager<T> GREY_25_PERCENT();
		IStyleManager<T> ROSE();
		IStyleManager<T> LIGHT_YELLOW();
		IStyleManager<T> LIGHT_GREEN();
		IStyleManager<T> LIGHT_TURQUOISE();
		IStyleManager<T> PALE_BLUE();
		IStyleManager<T> LAVENDER();
		IStyleManager<T> WHITE();
		IStyleManager<T> CORNFLOWER_BLUE();
		IStyleManager<T> LEMON_CHIFFON();
		IStyleManager<T> MAROON();
		IStyleManager<T> ORCHID();
		IStyleManager<T> CORAL();
		IStyleManager<T> ROYAL_BLUE();
		IStyleManager<T> LIGHT_CORNFLOWER_BLUE();
		IStyleManager<T> TAN();
		IStyleManager<T> name(String name);
	}
	
	public interface IAlignStyle<T> extends IStyleBase<T> {
		IStyleManager<T> ALIGN_GENERAL();
		IStyleManager<T> ALIGN_LEFT();
		IStyleManager<T> ALIGN_CENTER();
		IStyleManager<T> ALIGN_RIGHT();
		IStyleManager<T> ALIGN_FILL();
		IStyleManager<T> ALIGN_JUSTIFY();
		IStyleManager<T> ALIGN_CENTER_SELECTION();
	}
	
	public interface IVAlignStyle<T> extends IStyleBase<T> {
		IStyleManager<T> VERTICAL_TOP();
		IStyleManager<T> VERTICAL_CENTER();
		IStyleManager<T> VERTICAL_BOTTOM();
		IStyleManager<T> VERTICAL_JUSTIFY();
	}
	
	public interface IBorderType<T> extends IStyleBase<T> {
		IBorderKind<T> top();
		IBorderKind<T> left();
		IBorderKind<T> right();
		IBorderKind<T> bottom();
		IBorderKind<T> all();
	}
	
	public interface IBorderKind<T> extends IStyleBase<T> {
		IBorderColor<T> color();
		IBorderStyle<T> type();
	}
	
	public interface IBorderStyle<T> extends IStyleBase<T> {
		IStyleManager<T> BORDER_NONE();
		IStyleManager<T> BORDER_THIN();
		IStyleManager<T> BORDER_MEDIUM();
		IStyleManager<T> BORDER_DASHED();
		IStyleManager<T> BORDER_DOTTED();
		IStyleManager<T> BORDER_THICK();
		IStyleManager<T> BORDER_DOUBLE();
		IStyleManager<T> BORDER_HAIR();
		IStyleManager<T> BORDER_MEDIUM_DASHED();
		IStyleManager<T> BORDER_DASH_DOT();
		IStyleManager<T> BORDER_MEDIUM_DASH_DOT();
		IStyleManager<T> BORDER_DASH_DOT_DOT();
		IStyleManager<T> BORDER_MEDIUM_DASH_DOT_DOT();
		IStyleManager<T> BORDER_SLANTED_DASH_DOT();
		IStyleManager<T> clear();
	}
	
	public interface IBorderColor<T> extends IColor<T> {}
	
	public interface IFontType<T> extends IStyleBase<T> {
		IStyleManager<T> name(String name);
		IFontColor<T> color();
		IStyleManager<T> italic(boolean value);
		IStyleManager<T> NORMAL();
		IStyleManager<T> BOLD();
		IStyleManager<T> fontHeight(short value);
		IStyleManager<T> fontHeight(int value);
		IStyleManager<T> points(short value);
		IStyleManager<T> points(int value);
		IStyleManager<T> strikeout(boolean value);
		IFontOffset<T> typeOffset();
		IFontUnderline<T> underline();
	}
	
	public interface IFontOffset<T> extends IStyleBase<T> {
		IStyleManager<T> SS_NONE();
		IStyleManager<T> SS_SUPER();
		IStyleManager<T> SS_SUB();
	}
	
	public interface IFontUnderline<T> extends IStyleBase<T> {
		IStyleManager<T> U_NONE();
		IStyleManager<T> U_SINGLE();
		IStyleManager<T> U_DOUBLE();
		IStyleManager<T> U_SINGLE_ACCOUNTING();
		IStyleManager<T> U_DOUBLE_ACCOUNTING();
	}
	public interface IFontColor<T> extends IColor<T> {}
	
	
	public interface IFillType<T> {
		
		IForegroundColor<T> foregroundColor();
		IBackgroundColor<T> backgroundColor();
		IFillPattern<T> pattern();
	}
	
	public interface IForegroundColor<T> extends IColor<T> {}
	public interface IBackgroundColor<T> extends IColor<T> {}
	
	public interface IFillPattern<T> extends IStyleBase<T> {
		IStyleManager<T> NO_FILL();
		IStyleManager<T> SOLID_FOREGROUND();
		IStyleManager<T> FINE_DOTS();
		IStyleManager<T> ALT_BARS();
		IStyleManager<T> SPARSE_DOTS();
		IStyleManager<T> THICK_HORZ_BANDS();
		IStyleManager<T> THICK_VERT_BANDS();
		IStyleManager<T> THICK_BACKWARD_DIAG();
		IStyleManager<T> THICK_FORWARD_DIAG();
		IStyleManager<T> BIG_SPOTS();
		IStyleManager<T> BRICKS();
		IStyleManager<T> THIN_HORZ_BANDS();
		IStyleManager<T> THIN_VERT_BANDS();
		IStyleManager<T> THIN_BACKWARD_DIAG();
		IStyleManager<T> THIN_FORWARD_DIAG();
		IStyleManager<T> SQUARES();
		IStyleManager<T> DIAMONDS();
	}
	
	public interface IStyleBase<T> {
		IStyleManager<T> getStyle();
	}
}
