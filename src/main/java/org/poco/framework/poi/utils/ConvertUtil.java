package org.poco.framework.poi.utils;

import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;
import org.poco.framework.poi.converter.impl.BooleanConverter;
import org.poco.framework.poi.converter.impl.CalendarConverter;
import org.poco.framework.poi.converter.impl.DateConverter;
import org.poco.framework.poi.converter.impl.PoiNumberConverter;
import org.poco.framework.poi.converter.impl.StringConverter;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;

public class ConvertUtil {
	
	public static Object convertPrimitive(Object value) {
		if (!value.getClass().isPrimitive()) {
			return value;
		}
		else if (value.getClass() == double.class)
		{
			return new Double(value.toString());
		}
		else if (value.getClass() == float.class) {
			return new Float(value.toString());
		}
		else if (value.getClass() == int.class) {
			return new Integer(value.toString());
		}
		else if (value.getClass() == long.class) {
			return new Long(value.toString());
		}
		else if (value.getClass() == byte.class) {
			return new Byte(value.toString());
		}
		else if (value.getClass() == short.class){
			return new Short(value.toString());
		}
		else if (value.getClass() == boolean.class) {
			return (Boolean)value;
		}
		return null;
	}
	
	public static Object getValue(HSSFCell cell) {
		Object result = null;
		//タイプ分岐
		switch (cell.getCellType())
		{
			case Cell.CELL_TYPE_BLANK:
				result = "";
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				result = cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_ERROR:
				result = cell.getErrorCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				result = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				double num = cell.getNumericCellValue();
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					///日付型
					result = HSSFDateUtil.getJavaDate(num);
				} else {
					// 数値型
					result = cell.getNumericCellValue();
				}
				break;
			case Cell.CELL_TYPE_STRING:
				result = cell.getStringCellValue();
				break;
		}
		return result;
	}

	public static IConverter getConverter(Object param, String datePattern) {
		
		if (param != null) {

			Object value = convertPrimitive(param);
			
			if (value instanceof String) {
				return new StringConverter((String)value, datePattern);
			}
			else if (value instanceof Number)
			{
				return new PoiNumberConverter((Number)value, datePattern);
			}
			else if (value instanceof Date) {
				return new DateConverter((Date)value, datePattern);
			}
			else if (value instanceof Calendar) {
				return new CalendarConverter((Calendar)value, datePattern);
			}
			else if (value instanceof Boolean) {
				return new BooleanConverter((Boolean)value, datePattern);
			}
			else {
				throw new java.lang.ClassCastException("not support.");
			}
		}
		return null;
	}
	
	public static IConverter getConverter(Object param) {
		
		if (param != null) {

			Object value = convertPrimitive(param);
			
			if (value instanceof String) {
				return new StringConverter((String)value);
			}
			else if (value instanceof Number)
			{
				return new PoiNumberConverter((Number)value);
			}
			else if (value instanceof Date) {
				return new DateConverter((Date)value);
			}
			else if (value instanceof Calendar) {
				return new CalendarConverter((Calendar)value);
			}
			else if (value instanceof Boolean) {
				return new BooleanConverter((Boolean)value);
			}
			else {
				throw new java.lang.ClassCastException("not support.");
			}
		}
		return null;
	}
	
	public static Object convert(IConverter conv, Class<?> toClass){
		if (toClass == String.class) {
			return conv.toString();
		}
		else if(toClass == Long.class || toClass == long.class) {
			return conv.toLong();
		}
		else if(toClass == Integer.class || toClass == int.class) {
			return conv.toInteger();
		}
		else if(toClass == Double.class || toClass == double.class ) {
			return conv.toDouble();
		}
		else if(toClass == Calendar.class ) {
			return conv.toCalender();
		}
		else if(toClass == Date.class ) {
			return conv.toDate();
		}
		else if (toClass == Float.class || toClass == float.class) {
			return conv.toFloat();
		}
		else if (toClass == Boolean.class || toClass == boolean.class) {
			return conv.toBoolean();
		}
		return null;
	}
	
}
