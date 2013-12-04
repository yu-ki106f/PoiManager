/**
 * 
 */
package org.poco.framework.poi.converter.impl;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;

/**
 * @author yu-ki106f
 *
 */
public class PoiNumberConverter extends NumberConverter {
	
	public PoiNumberConverter(Number value) {
		super(value);
	}
	public PoiNumberConverter(Number value, String datePattern) {
		super(value, datePattern);
	}

	@Override
	public Calendar toCalender() {
		Date d = toDate();
		if (d == null) return null;
		Calendar result = Calendar.getInstance();
		result.setTime(d);
		return result;
	}

	@Override
	public Date toDate() {
		return value == null ? null : HSSFDateUtil.getJavaDate(toDouble());
	}
}
