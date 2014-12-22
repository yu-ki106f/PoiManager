/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.BoldWeight;
import org.poco.framework.poi.constants.PoiConstants.FLColor;
import org.poco.framework.poi.constants.PoiConstants.FontOffset;
import org.poco.framework.poi.constants.PoiConstants.FontUnderline;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IFontColor;
import org.poco.framework.poi.managers.IStyleManager.IFontOffset;
import org.poco.framework.poi.managers.IStyleManager.IFontType;
import org.poco.framework.poi.managers.IStyleManager.IFontUnderline;

/**
 * @author yu-ki
 *
 */
public class FontStyleCreator<T> extends StyleCreatorImpl<T> {

	public FontStyleCreator(IStyleManager<T> style, PoiStyleDto dto) {
		super(style, dto);
	}

	public IFontType<T> createFontType()
	{
		return new FontTypeImpl<T>(this);
	}
	
	public static class FontTypeImpl<T> extends AbstractFontTypeImpl<T> implements IFontType<T>
	{

		private StyleCreatorImpl<T> creator;
		
		public FontTypeImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}

	
		protected void update(BoldWeight boldWeight) {
			creator.getStyleDto().getFont().boldweight = boldWeight;
		}

	
		public IFontColor<T> color() {
			return new FontColorImpl<T>(this.creator);
		}

	
		public IStyleManager<T> fontHeight(short value) {
			creator.getStyleDto().getFont().fontHeight = value;
			creator.getStyleDto().getFont().fontHeightInPoints = (short)(value / 20);
			return getStyle();
		}

	
		public IStyleManager<T> italic(boolean value) {
			creator.getStyleDto().getFont().italic = value;
			return getStyle();
		}

	
		public IStyleManager<T> name(String name) {
			creator.getStyleDto().getFont().fontName = name;
			return getStyle();
		}

	
		public IStyleManager<T> points(short value) {
			creator.getStyleDto().getFont().fontHeight = (short)(value * 20);
			creator.getStyleDto().getFont().fontHeightInPoints = value;
			return getStyle();
		}

	
		public IStyleManager<T> strikeout(boolean value) {
			creator.getStyleDto().getFont().strikeout = value;
			return getStyle();
		}

	
		public IFontOffset<T> typeOffset() {
			return new FontOffsetImpl<T>(this.creator);
		}

	
		public IFontUnderline<T> underline() {
			return new FontUnderlineImpl<T>(this.creator);
		}

	
		public IStyleManager<T> fontHeight(int value) {
			if (Short.MAX_VALUE < value || Short.MIN_VALUE > value) {
				throw new java.lang.ClassCastException("short over parametor");
			}
			return fontHeight((short)value);
		}

	
		public IStyleManager<T> points(int value) {
			if (Short.MAX_VALUE < value || Short.MIN_VALUE > value) {
				throw new java.lang.ClassCastException("short over parametor");
			}
			return points((short)value);
		}
		
	}
	
	public static class FontUnderlineImpl<T> extends AbstractFontUnderlineImpl<T> implements IFontUnderline<T>
	{

		private StyleCreatorImpl<T> creator;
		
		public FontUnderlineImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
	
		protected void update(FontUnderline underline) {
			creator.getStyleDto().getFont().underline = underline;
			
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}

	}

	public static class FontColorImpl<T> extends AbstractColorImpl<T> implements IFontColor<T>
	{
		private StyleCreatorImpl<T> creator;
		
		public FontColorImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(FLColor color) {
			creator.getStyleDto().getFont().color = color;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
	}
	
	public static class FontOffsetImpl<T> extends AbstractFontOffsetImpl<T> implements IFontOffset<T>
	{
		private StyleCreatorImpl<T> creator;
		
		public FontOffsetImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(FontOffset offset) {
			creator.getStyleDto().getFont().typeOffset = offset;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
		
	}
}
