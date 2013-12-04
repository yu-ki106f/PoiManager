package org.poco.framework.poi.creator.image.impl;

import org.poco.framework.poi.constants.PoiConstants.AnchorType;
import org.poco.framework.poi.creator.image.IImageCreator;
import org.poco.framework.poi.creator.image.IImageCreator.IAnchorType;

public abstract class AbstractAnchorType<T> implements IAnchorType<T> {
	
	abstract protected void update(AnchorType type);
	
	public IImageCreator<T> DONT_MOVE_AND_RESIZE() {
		update(AnchorType.DONT_MOVE_AND_RESIZE);
		return getCreator();
	}

	public IImageCreator<T> MOVE_AND_RESIZE() {
		update(AnchorType.MOVE_AND_RESIZE);
		return getCreator();
	}

	public IImageCreator<T> MOVE_DONT_RESIZE() {
		update(AnchorType.MOVE_DONT_RESIZE);
		return getCreator();
	}

}
