package org.poco.framework.poi.converter;

import java.util.Calendar;
import java.util.Date;

public interface IValueConverter {
	Date getDate();
	Calendar getCalendar();
	String getString();
	Double getDouble();
	Float getFloat();
	Long getLong();
	Integer getInteger();
	Byte getByte();
	Boolean getBoolean();
	Object getOriginalValue();
}
