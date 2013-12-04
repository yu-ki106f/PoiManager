package org.poco.framework.poi.dto;

public class PoiRect {

	public Integer top;
	public Integer left;
	public Integer right;
	public Integer bottom;
	
	public PoiRect(PoiPosition... value) {
		if (value.length == 1) {
			top =  value[0].y;
			bottom = value[0].y;
			right = value[0].x;
			left = value[0].x;
		}
		else if (value.length >= 2) {
			top = value[0].y;
			left = value[0].x;
			bottom = value[1].y;
			right = value[1].x;
		}
	}
	
	public PoiRect(Integer top, Integer left, Integer right, Integer bottom) {
		this.top =  top;
		this.left = left;
		this.right = right;
		this.bottom = bottom;
	}

	public boolean equals(PoiRect rect) {
		
		return (this.top == rect.top 
				&& this.left == rect.left 
				&& this.right == rect.right
				&& this.bottom == rect.bottom);
	}
}
