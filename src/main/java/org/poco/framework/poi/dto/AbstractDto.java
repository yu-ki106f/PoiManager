package org.poco.framework.poi.dto;

import java.lang.reflect.Field;

public abstract class AbstractDto {
	
	protected boolean isChange() {
		for (Field f : this.getClass().getFields()) {
			try {
				if (f.get(this) != null ) {
					return true;
				}
			} catch (Exception e) {
				return false;
			}
		}
		return false;
	}
	
	/**
	 * 比較
	 * @param dto
	 * @return
	 */
	public boolean equals(AbstractDto dto) {
		for (Field f : this.getClass().getFields()) {
			try {
				if (f.get(this) == null && f.get(dto) == null) {
				}
				else if (!f.get(this).equals(f.get(dto)) ) {
					return false;
				}
			} catch (Exception e) {
				return false;
			}
		}
		return true;
	}
}
