package org.poco.framework.poi.converter.impl;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;

public class StringConverter extends AConverter<String> implements IConverter {

	public StringConverter(String value) {
		super(value);
	}

	public StringConverter(String value,String datePattern) {
		super(value,datePattern);
	}


	public Byte toByte() {
		return new Byte(value);
	}


	public Calendar toCalender() {
		if (value == null) return null;

		Calendar result = null;
		Date date = toDate();
		if (date != null) {
			result = Calendar.getInstance();
			result.setTime(date);
		}
		return result;
	}


	public Date toDate() {
		if (value == null) return null;
		
		SimpleDateFormat df = new SimpleDateFormat(datePattern);
		Date result = null;
		try {
			result = df.parse(value);
		} catch (ParseException e) {
		}
		return result;
	}


	public Double toDouble() {
		return value == null ? null : new Double(value);
	}


	public Float toFloat() {
		return value == null ? null : new Float(value);
	}


	public Integer toInteger() {
		return value == null ? null : new Integer(value);
	}


	public Long toLong() {
		return value == null ? null : new Long(value);
	}


	public Short toShort() {
		return value == null ? null : new Short(value);
	}
	

	public String toString() {
		return value;
	}
}
