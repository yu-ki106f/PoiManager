package org.poco.framework.poi.creator.image;

import org.poco.framework.poi.creator.image.impl.ImageCreatorImpl.PoiImageDto;
import org.poco.framework.poi.dto.PoiRect;

public interface IImageCreator<T> {
	T drow();
	
	IImageCreator<T> padding(int left, int top, int right, int bottom);

	IImageCreator<T> padding(PoiRect rect);
	
	IImageCreator<T> padding(int all);

	IAnchorType<T> anchoType();
	
	PoiImageDto getImageDto();
	
	public interface IAnchorType<T> extends IImageBase<T> {
		
		IImageCreator<T> MOVE_AND_RESIZE();

		IImageCreator<T> MOVE_DONT_RESIZE();
		
		IImageCreator<T> DONT_MOVE_AND_RESIZE();
	}
	
	public interface IImageBase<T> {
		IImageCreator<T> getCreator();
	}
}
