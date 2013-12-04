package org.poco.framework.poi.converter.impl;

import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;

public class CalendarConverter extends AConverter<Calendar> implements IConverter {

	private DateConverter conv = null;
	
	public CalendarConverter(Calendar value) {
		super(value);
		conv = new DateConverter(value.getTime());
	}

	public CalendarConverter(Calendar value, String datePattern) {
		super(value, datePattern);
		conv = new DateConverter(value.getTime(),datePattern);
	}


	public void setDatePattern(String datePattern) {
		super.datePattern = datePattern;
		conv.setDatePattern(datePattern);
	}
	

	public Byte toByte() {
		return conv.toByte();
	}


	public Calendar toCalender() {
		return value;
	}


	public Date toDate() {
		return value == null ? null : value.getTime();
	}


	public Double toDouble() {
		return conv.toDouble();
	}


	public Float toFloat() {
		return conv.toFloat();
	}


	public Integer toInteger() {
		return conv.toInteger();
	}


	public Long toLong() {
		return conv.toLong();
	}


	public Short toShort() {
		return conv.toShort();
	}
	

	public String toString() {
		return conv.toString();
	}

}
