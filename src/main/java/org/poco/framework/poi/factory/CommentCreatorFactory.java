package org.poco.framework.poi.factory;

import org.poco.framework.poi.creator.comment.ICommentCreator;
import org.poco.framework.poi.creator.comment.impl.CommentCreator;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiRange;

public class CommentCreatorFactory {

	public static ICommentCreator<IPoiRange> create(IPoiRange value, String comment) {
		CommentCreator<IPoiRange> result = new CommentCreator<IPoiRange>(comment);
		result.parent = value;
		return result;
	}
	
	public static ICommentCreator<IPoiCell> create(IPoiCell value, String comment) {
		CommentCreator<IPoiCell> result = new CommentCreator<IPoiCell>(comment);
		result.parent = value;
		return result;
	}
}
