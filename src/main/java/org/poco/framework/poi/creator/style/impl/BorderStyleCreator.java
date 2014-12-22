package org.poco.framework.poi.creator.style.impl;

import org.poco.framework.poi.constants.PoiConstants.Border;
import org.poco.framework.poi.constants.PoiConstants.BorderPosition;
import org.poco.framework.poi.constants.PoiConstants.FLColor;
import org.poco.framework.poi.creator.style.IStyleCreator;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.managers.IStyleManager.IBorderColor;
import org.poco.framework.poi.managers.IStyleManager.IBorderKind;
import org.poco.framework.poi.managers.IStyleManager.IBorderStyle;
import org.poco.framework.poi.managers.IStyleManager.IBorderType;

public class BorderStyleCreator<T> extends StyleCreatorImpl<T> {

	public BorderStyleCreator(IStyleManager<T> style, PoiStyleDto dto) {
		super(style, dto);
	}

	/**
	 * BorderTypeを作成します
	 * @return
	 */
	public IBorderType<T> createBorderType() {
		return new BorderTypeImpl<T>(this);
	}
	
	/**
	 * IBorderTypeの実装
	 * @author yu-ki
	 *
	 */
	public static class BorderTypeImpl<T> implements IBorderType<T>
	{
		private IStyleCreator<T> creator;
		
		public BorderTypeImpl(IStyleCreator<T> creator) {
			this.creator = creator;
		}
		
	
		public IBorderKind<T> bottom() {
			return new BoaderKindImpl<T>(creator, BorderPosition.bottom);
		}

	
		public IBorderKind<T> left() {
			return new BoaderKindImpl<T>(creator,BorderPosition.left);
		}

	
		public IBorderKind<T> right() {
			return new BoaderKindImpl<T>(creator,BorderPosition.right);
		}

	
		public IBorderKind<T> top() {
			return new BoaderKindImpl<T>(creator,BorderPosition.top);
		}

	
		public IBorderKind<T> all() {
			return new BoaderKindImpl<T>(creator,BorderPosition.all);
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}
	}

	/**
	 * 中継インターフェース
	 * @author yu-ki
	 *
	 */
	public interface IKindTyep<T> {
		
		public BorderPosition getType();

		public IStyleCreator<T> getCreator();
	}
	
	/**
	 * ボーダー種類の実装
	 * @author yu-ki
	 *
	 */
	public static class BoaderKindImpl<T> implements IBorderKind<T>, IKindTyep<T>
	{
		private IStyleCreator<T> creator;
		
		private BorderPosition type;
		
		public BoaderKindImpl(IStyleCreator<T> creator, BorderPosition type) {
			this.creator = creator;
			this.type = type;
		}

		public BorderPosition getType() {
			return type;
		}

		public void setType(BorderPosition type) {
			this.type = type;
		}

		public IStyleCreator<T> getCreator() {
			return creator;
		}

	
		public IBorderColor<T> color() {
			return new BorderColorImpl<T>(this);
		}

	
		public IBorderStyle<T> type() {
			return new BorderStyleImpl<T>(this);
		}

	
		public IStyleManager<T> getStyle() {
			return creator.getStyle();
		}

	}
	
	/**
	 * IBorderColorの実装
	 * @author yu-ki
	 *
	 */
	public static class BorderColorImpl<T> extends AbstractColorImpl<T> implements IBorderColor<T>
	{
		private IKindTyep<T> type;
		
		public BorderColorImpl(IKindTyep<T> type) {
			this.type = type;
		}

	
		protected void update(FLColor color) {
			
			switch(type.getType()) {
			case top:
				this.type.getCreator().getStyleDto().topBorderColor = color;
				break;
			case bottom:
				this.type.getCreator().getStyleDto().bottomBorderColor = color;
				break;
			case left:
				this.type.getCreator().getStyleDto().leftBorderColor = color;
				break;
			case right:
				this.type.getCreator().getStyleDto().rightBorderColor = color;
				break;
			case all:
				this.type.getCreator().getStyleDto().topBorderColor = color;
				this.type.getCreator().getStyleDto().bottomBorderColor = color;
				this.type.getCreator().getStyleDto().leftBorderColor = color;
				this.type.getCreator().getStyleDto().rightBorderColor = color;
				break;
			}
		}

	
		public IStyleManager<T> getStyle() {
			return type.getCreator().getStyle();
		}

	}

	/**
	 * 
	 * @author yu-ki
	 *
	 */
	public static class BorderStyleImpl<T> extends AbstractBorderImpl<T> implements IBorderStyle<T>
	{
		private IKindTyep<T> type;
		
		public BorderStyleImpl(IKindTyep<T> type) {
			this.type = type;
		}

	
		protected void update(Border border) {
			switch(type.getType()) {
			case top:
				this.type.getCreator().getStyleDto().borderTop= border;
				break;
			case bottom:
				this.type.getCreator().getStyleDto().borderBottom = border;
				break;
			case left:
				this.type.getCreator().getStyleDto().borderLeft = border;
				break;
			case right:
				this.type.getCreator().getStyleDto().borderRight = border;
				break;
			case all:
				this.type.getCreator().getStyleDto().borderTop= border;
				this.type.getCreator().getStyleDto().borderBottom = border;
				this.type.getCreator().getStyleDto().borderLeft = border;
				this.type.getCreator().getStyleDto().borderRight = border;
				break;
			}
		}

	
		public IStyleManager<T> getStyle() {
			return type.getCreator().getStyle();
		}

	}
}
