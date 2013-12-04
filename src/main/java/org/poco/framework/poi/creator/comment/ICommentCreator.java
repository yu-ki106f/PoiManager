package org.poco.framework.poi.creator.comment;

import org.poco.framework.poi.dto.PoiRect;

public interface ICommentCreator<T> {
	T write();
	
	ICommentCreator<T> padding(int left, int top, int right, int bottom);
	
	ICommentCreator<T> padding(PoiRect rect);
	
	ICommentCreator<T> padding(int all);

	ICommentCreator<T> startPosDiff(int xStart, int yStart, int xEnd, int yEnd);

	ICommentCreator<T> comment(String value);
	
	ICommentCreator<T> createUser(String value);
}
