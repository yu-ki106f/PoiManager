package org.poco.framework.poi.factory;

import java.io.File;

import org.poco.framework.poi.creator.image.IImageCreator;
import org.poco.framework.poi.creator.image.impl.ImageCreatorImpl;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiRange;

public class ImageCreatorFactory {
	
	public static IImageCreator<IPoiRange> create(IPoiRange value, File f) {
		ImageCreatorImpl<IPoiRange> result = new ImageCreatorImpl<IPoiRange>(f);
		result.parent = value;
		return result;
	}

	public static IImageCreator<IPoiCell> create(IPoiCell value, File f) {
		ImageCreatorImpl<IPoiCell> result = new ImageCreatorImpl<IPoiCell>(f);
		result.parent = value;
		return result;
	}
	
	public static IImageCreator<IPoiRange> create(IPoiRange value, byte[] bytes, String fileName) {
		ImageCreatorImpl<IPoiRange> result = new ImageCreatorImpl<IPoiRange>(bytes, fileName);
		result.parent = value;
		return result;
	}

	public static IImageCreator<IPoiCell> create(IPoiCell value, byte[] bytes, String fileName) {
		ImageCreatorImpl<IPoiCell> result = new ImageCreatorImpl<IPoiCell>(bytes, fileName);
		result.parent = value;
		return result;
	}
}
