package org.poco.framework.poi.managers;

import org.poco.framework.poi.exception.PoiException;
import org.poco.framework.poi.managers.IPoiManager.IPoiBook;

public interface IReadWriter {
	public IReadWriter loadFromXml(String fileName) throws PoiException;
	public IPoiBook write(Object data) throws PoiException;
	public <T> T read(Class<T> baseClass) throws PoiException;
	public IReadWriter setMaxRows(Integer value);
	public IReadWriter setStyleCopy(boolean value);
}
