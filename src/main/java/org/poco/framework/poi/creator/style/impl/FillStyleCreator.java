package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.Color;
import org.poco.framework.poi.constants.PoiConstants.FillPattern;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IBackgroundColor;
import org.poco.framework.poi.managers.IStyleManager.IFillPattern;
import org.poco.framework.poi.managers.IStyleManager.IFillType;
import org.poco.framework.poi.managers.IStyleManager.IForegroundColor;

/**
 * @author yu-ki106f
 *
 */
public class FillStyleCreator<T> extends StyleCreatorImpl<T> {

	public FillStyleCreator(IStyleManager<T> style, PoiStyleDto dto) {
		super(style, dto);
	}

	public IFillType<T> createFillType() {
		return new FillTypeImpl<T>(this);
	}
	
	public static class FillTypeImpl<T> implements IFillType<T>
	{
		private StyleCreatorImpl<T> creator;
		
		public FillTypeImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		public IForegroundColor<T> foregroundColor() {
			return new ForegroundColorImpl<T>(creator);
		}

	
		public IFillPattern<T> pattern() {
			return new FillPatternImpl<T>(creator);
		}

	
		public IBackgroundColor<T> backgroundColor() {
			return new BackgroundColorImpl<T>(creator);
		}
	}
	
	public static class ForegroundColorImpl<T> extends AbstractColorImpl<T> implements IForegroundColor<T>
	{
		private StyleCreatorImpl<T> creator;
		
		public ForegroundColorImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(Color color) {
			this.creator.getStyleDto().fillForegroundColor = color;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
	}
	
	public static class BackgroundColorImpl<T> extends AbstractColorImpl<T> implements IBackgroundColor<T>
	{
		private StyleCreatorImpl<T> creator;
		
		public BackgroundColorImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(Color color) {
			this.creator.getStyleDto().fillBackgroundColor = color;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
	}
	
	public static class FillPatternImpl<T> extends AbstractFillPatternImpl<T> implements IFillPattern<T>
	{
		private StyleCreatorImpl<T> creator;
		
		public FillPatternImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(FillPattern fill) {
			creator.getStyleDto().fillPattern = fill;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
		
	}
}
