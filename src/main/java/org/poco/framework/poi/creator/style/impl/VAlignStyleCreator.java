/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.VAlign;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IVAlignStyle;

/**
 * @author yu-ki106f
 *
 */
public class VAlignStyleCreator<T> extends StyleCreatorImpl<T> {

	public VAlignStyleCreator(IStyleManager<T> style, PoiStyleDto dto) {
		super(style, dto);
	}

	/**
	 * AlignStyleを作成します。
	 * @return
	 */
	public IVAlignStyle<T> createVAlignStyle() {
		return new VAlignStyleImpl<T>(this);
	}
	
	/**
	 * IAlignStyleの実装
	 * @author yu-ki106f
	 *
	 */
	public static class VAlignStyleImpl<T> extends AbstractVAlignImpl<T> implements IVAlignStyle<T>
	{
		private StyleCreatorImpl<T> creator;
		
		/**
		 * コンストラクタ
		 * @param creator
		 */
		public VAlignStyleImpl(StyleCreatorImpl<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(VAlign valign) {
			this.creator.getStyleDto().valign = valign;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
		
	}
}
