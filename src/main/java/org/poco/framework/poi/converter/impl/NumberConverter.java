package org.poco.framework.poi.converter.impl;

import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;

public class NumberConverter extends AConverter<Number> implements IConverter {

	public NumberConverter(Number value) {
		super(value);
	}

	public NumberConverter(Number value, String datePattern) {
		super(value, datePattern);
	}


	public Byte toByte() {
		return value == null ? null : value.byteValue();
	}


	public Calendar toCalender() {
		if (value == null) return null;
		Calendar cal = Calendar.getInstance();
		cal.setTimeInMillis(value.longValue());
		return cal;
	}


	public Date toDate() {
		return value == null ? null : new Date(value.longValue());
	}


	public Double toDouble() {
		return value == null ? null : value.doubleValue();
	}


	public Float toFloat() {
		return value == null ? null : value.floatValue();
	}


	public Integer toInteger() {
		return value == null ? null : value.intValue();
	}


	public Long toLong() {
		return value == null ? null : value.longValue();
	}


	public Short toShort() {
		return value == null ? null : value.shortValue();
	}
	

	public String toString() {
		return value == null ? null : value.toString();
	}
}
