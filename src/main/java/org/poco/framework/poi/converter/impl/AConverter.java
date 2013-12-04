package org.poco.framework.poi.converter.impl;

import org.poco.framework.poi.converter.IConverter;

public abstract class AConverter<T> implements IConverter {

	protected T value = null;
	protected String datePattern = "yyyy/MM/dd HH:MM:ss";
	
	public AConverter(T value) {
		this.value = value;
	}

	public AConverter(T value, String datePattern) {
		this.value = value;
		this.datePattern = datePattern;
	}
	

	public void setDatePattern(String datePattern) {
		this.datePattern = datePattern;
	}
	
	public Boolean toBoolean() {
		return value == null ? false : true;
	}
}
