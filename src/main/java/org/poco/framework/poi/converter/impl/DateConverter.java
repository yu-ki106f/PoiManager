package org.poco.framework.poi.converter.impl;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;

public class DateConverter extends AConverter<Date> implements IConverter {

	public DateConverter(Date value) {
		super(value);
	}

	public DateConverter(Date value, String datePattern) {
		super(value, datePattern);
	}


	public Byte toByte() {
		return toDouble().byteValue();
	}


	public Calendar toCalender() {
		if (value == null) return null;
		Calendar cal = Calendar.getInstance();
		cal.setTime(value);
		return cal;
	}


	public Date toDate() {
		return value == null ? null : value;
	}


	public Double toDouble() {
		return value == null ? null : Double.valueOf(value.getTime());
	}


	public Float toFloat() {
		return toDouble().floatValue();
	}


	public Integer toInteger() {
		return toDouble().intValue();
	}


	public Long toLong() {
		return toDouble().longValue();
	}


	public Short toShort() {
		return toDouble().shortValue();
	}
	

	public String toString() {
		if (value == null) return null;
		SimpleDateFormat df = new SimpleDateFormat(datePattern);
		return df.format(value);
	}

}
