/**
 * 
 */
package org.poco.framework.poi.converter.impl;

import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;

/**
 * @author yu-ki106f
 *
 */
public class BooleanConverter extends AConverter<Boolean> implements IConverter {

	public BooleanConverter(Boolean value) {
		super(value);
	}

	public BooleanConverter(Boolean value, String datePattern) {
		super(value, datePattern);
	}


	public Byte toByte() {
		return toShort().byteValue();
	}


	public Calendar toCalender() {
		return null;
	}


	public Date toDate() {
		return null;
	}


	public Double toDouble() {
		return toShort().doubleValue();
	}


	public Float toFloat() {
		return toShort().floatValue();
	}


	public Integer toInteger() {
		return toShort().intValue();
	}


	public Long toLong() {
		return toShort().longValue();
	}


	public Short toShort() {
		return value == null ? null : value ? (short)1 : (short)0;
	}
	
	public String toString() {
		return value.toString();
	}

}
