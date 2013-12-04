package org.poco.framework.poi.converter.impl;

import java.util.Calendar;
import java.util.Date;

import org.poco.framework.poi.converter.IConverter;
import org.poco.framework.poi.converter.IValueConverter;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.utils.ConvertUtil;

import org.apache.poi.ss.usermodel.Cell;

public class ValueConverter implements IValueConverter {
	
	private IPoiCell cell = null;
	
	public ValueConverter(IPoiCell cell) {
		this.cell = cell;
	}
	
	public Boolean getBoolean() {
		return _getValue();
	}

	public Byte getByte() {
		return _getValue();
	}

	public Calendar getCalendar() {
		return _getValue();
	}

	public Double getDouble() {
		return _getValue();
	}

	public Float getFloat() {
		return _getValue();
	}

	public Integer getInteger() {
		return _getValue();
	}

	public Long getLong() {
		return _getValue();
	}

	public String getString() {
		return _getValue();
	}
	
	public Date getDate()
	{
		return _getValue();
	}

	@SuppressWarnings("unchecked")
	private <T> T _getValue(T ...dummy ) {
		Object result = getOriginalValue();
		
		Class<?> clazz = dummy.getClass().getComponentType();
		
		IConverter conv = ConvertUtil.getConverter(result);
		
		return (T)ConvertUtil.convert(conv, clazz);
	}
	
	public Object getOriginalValue() {
		return ConvertUtil.getValue(cell.getOrgCell());
	}
	
	public <T> void setValue(Object param, Class<T> clazz) {
		IConverter conv = ConvertUtil.getConverter(param);
		//コンバート
		Object value = ConvertUtil.convert(conv, clazz) ;		

		setValue(value);
	}

	public void setFormura(String param) {
		cell.getOrgCell().setCellFormula(param);
		//cell.getOrgCell().setCellType(Cell.CELL_TYPE_FORMULA);
	}
	
	public void setError(Byte param) {
		cell.getOrgCell().setCellValue(param);
		cell.getOrgCell().setCellType(Cell.CELL_TYPE_ERROR);
	}

	public void setValue(Object param) {
		if (param == null) {
			cell.getOrgCell().setCellValue("");
			cell.getOrgCell().setCellType(Cell.CELL_TYPE_BLANK);
			return;
		}
		Object value = ConvertUtil.convertPrimitive(param);
		if (value instanceof String) {
			cell.getOrgCell().setCellValue((String)value);
			cell.getOrgCell().setCellType(Cell.CELL_TYPE_STRING);
		}
		else if (value instanceof Number) {
			cell.getOrgCell().setCellValue( new Double(value.toString()) );
			cell.getOrgCell().setCellType(Cell.CELL_TYPE_NUMERIC);
		}
		else if (value instanceof Date) {
			cell.getOrgCell().setCellValue( (Date)value );
			//TOODスタイル
			cell.getOrgCell().setCellType(Cell.CELL_TYPE_NUMERIC);
		}
		else if (value instanceof Calendar) {
			cell.getOrgCell().setCellValue( (Calendar)value );
			//TOODスタイル
			cell.getOrgCell().setCellType(Cell.CELL_TYPE_NUMERIC);
		}
		else if (value instanceof Boolean) {
			cell.getOrgCell().setCellValue( (Boolean)value );
			cell.getOrgCell().setCellType(Cell.CELL_TYPE_BOOLEAN);
		}
	}

}
