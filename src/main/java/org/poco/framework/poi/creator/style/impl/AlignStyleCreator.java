/**
 * 
 */
package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.Align;
import org.poco.framework.poi.creator.style.IStyleCreator;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IAlignStyle;

/**
 * @author yu-ki106f
 *
 */
public class AlignStyleCreator<T> extends StyleCreatorImpl<T> {

	public AlignStyleCreator(IStyleManager<T> style, PoiStyleDto dto) {
		super(style, dto);
	}

	/**
	 * AlignStyleを作成します。
	 * @return
	 */
	public IAlignStyle<T> createAlignStyle() {
		return new AlignStyleImpl<T>(this);
	}
	
	/**
	 * IAlignStyleの実装
	 * @author yu-ki106f
	 *
	 */
	public static class AlignStyleImpl<T> extends AbstractAlignImpl<T> implements IAlignStyle<T>
	{
		private IStyleCreator<T> creator;
		
		/**
		 * コンストラクタ
		 * @param creator
		 */
		public AlignStyleImpl(IStyleCreator<T> creator) {
			this.creator = creator;
		}
		
	
		protected void update(Align align) {
			this.creator.getStyleDto().align = align;
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
	}
}
