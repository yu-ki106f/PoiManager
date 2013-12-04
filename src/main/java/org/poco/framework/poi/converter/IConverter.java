package org.poco.framework.poi.converter;

import java.util.Calendar;
import java.util.Date;

public interface IConverter {
	String toString();
	Double toDouble();
	Float toFloat();
	Long toLong();
	Integer toInteger();
	Byte toByte();
	Short toShort();
	Date toDate();
	Calendar toCalender();
	Boolean toBoolean();
	void setDatePattern(String value);
}
