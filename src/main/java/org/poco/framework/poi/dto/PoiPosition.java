package org.poco.framework.poi.dto;


public class PoiPosition {
	public Integer x;
	public Integer y;
	
	public PoiPosition(Integer x, Integer y) {
		this.x = x;
		this.y = y;
	}
	
	public boolean equals(PoiPosition pos) {
		return (this.x == pos.x && this.y == pos.y);
	}
}
