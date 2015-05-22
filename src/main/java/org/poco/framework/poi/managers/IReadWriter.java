package org.poco.framework.poi.managers;

import org.poco.framework.poi.exception.PoiException;
import org.poco.framework.poi.managers.IPoiManager.IPoiBook;
import org.poco.framework.poi.managers.impl.ReadWriter.XMLBook;

public interface IReadWriter {
	public IReadWriter loadFromXml(String fileName) throws PoiException;
	public IPoiBook write(Object data) throws PoiException;
	public <T> T read(Class<T> baseClass) throws PoiException;
	public IReadWriter setMaxRows(Integer value);
	public IReadWriter setStyleCopy(boolean value);
	public XMLBook getTempleteXMLClone();
	public IPoiBook write(Object data, XMLBook templete) throws PoiException;
	public <T> T read(Class<T> baseClass, XMLBook templete) throws PoiException;
}
