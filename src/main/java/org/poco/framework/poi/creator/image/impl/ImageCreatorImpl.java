package org.poco.framework.poi.creator.image.impl;

import java.io.File;

import org.poco.framework.poi.constants.PoiConstants.AnchorType;
import org.poco.framework.poi.creator.image.IImageCreator;
import org.poco.framework.poi.dto.PoiPosition;
import org.poco.framework.poi.dto.PoiRect;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiRange;
import org.poco.framework.poi.utils.WorkbookUtil;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

public class ImageCreatorImpl<T> implements IImageCreator<T> {

	private File file = null;
	
	private byte bytes[];
	
	private int fileType;
	
	private PoiImageDto imageDto = null;
	
	public T parent = null;
	
	public ImageCreatorImpl(File f) {
		imageDto = new PoiImageDto();
		file = f;
		fileType = WorkbookUtil.getPictureType(file);
	}
	
	public ImageCreatorImpl(byte[] bytes, String suffix) {
		imageDto = new PoiImageDto();
		file = null;
		this.bytes = bytes;
		fileType = WorkbookUtil.getPictureType(WorkbookUtil.getSuffix(suffix));
	}
	
	public IAnchorType<T> anchoType() {
		return new AnchorTypeImpl<T>(this);
	}

	private PoiRect getCellRange() {
		if (parent instanceof IPoiRange) {
			return getCellRange((IPoiRange)parent);
		}
		return getCellRange((IPoiCell)parent);
	}
	
	private Sheet getSheet() {
		if (parent instanceof IPoiRange) {
			return getSheet((IPoiRange)parent);
		}
		return getSheet((IPoiCell)parent);
	}

	private Sheet getSheet(IPoiRange value) {
		return value.getOrgSheet();
	}
	
	private Sheet getSheet(IPoiCell value) {
		return value.getOrgSheet();
	}
	
	private PoiRect getCellRange(IPoiRange value) {
		PoiRect result = new PoiRect();
		int length = value.getCells().length;
		result.top = value.getCells()[0].y();
		result.left = value.getCells()[0].x();
		result.right = value.getCells()[length - 1].x();
		result.bottom = value.getCells()[length - 1].y();
		
		return result;
	}
	private PoiRect getCellRange(IPoiCell value) {
		return new PoiRect(new PoiPosition(value.x(),value.y()));
	}
	
	public T drow() {
		PoiRect rect = getCellRange();
		byte[] bytes = null;
		
		Drawing patriarch = null;
		try {
			//patriarch = getSheet().getDrawingPatriarch();
			if (patriarch == null) {
				patriarch = getSheet().createDrawingPatriarch();
			}
			
			ClientAnchor clientAnchor = null;
			if (WorkbookUtil.isXlsx(getSheet().getWorkbook()) ) {
				clientAnchor = new XSSFClientAnchor();
			}
			else {
				clientAnchor = new HSSFClientAnchor();
			}
			clientAnchor.setDx1(imageDto.padding.left);
			clientAnchor.setDy1(imageDto.padding.top);
			//マイナス値にする
			clientAnchor.setDx2(-imageDto.padding.right);
			clientAnchor.setDy2(-imageDto.padding.bottom);
			
			clientAnchor.setCol1(rect.left);
			clientAnchor.setRow1(rect.top);
			//1つ次のCellにしておき、パディングを-値とする
			clientAnchor.setCol2(rect.right+1);
			clientAnchor.setRow2(rect.bottom+1);
			//アンカータイプ
			clientAnchor.setAnchorType(imageDto.anchorType.value());
			//バイト配列
			if (file != null) {
				bytes = WorkbookUtil.getByte(file);
			}
			else {
				bytes = this.bytes;
			}
			
			int add = getSheet().getWorkbook().addPicture(bytes, fileType);
			patriarch.createPicture(clientAnchor, add);
		}
		catch(Exception e) {
			//TODO ERROR
			e.printStackTrace();
		}
		return parent;
	}
	
	public IImageCreator<T> padding(int left, int top, int right, int bottom) {
		imageDto.padding.left = left;
		imageDto.padding.top = top;
		imageDto.padding.right = right;
		imageDto.padding.bottom = bottom;
		return this;
	}

	public IImageCreator<T> padding(int all) {
		return padding(all,all,all,all);
	}

	public IImageCreator<T> padding(PoiRect rect) {
		imageDto.padding = rect;
		return this;
	}

	public PoiImageDto getImageDto() {
		return imageDto;
	}

	
	public static class AnchorTypeImpl<T> extends AbstractAnchorType<T> implements IAnchorType<T>
	{
		private IImageCreator<T> creator; 
		
		public AnchorTypeImpl(IImageCreator<T> creator) {
			this.creator = creator;
		}
		
		public IImageCreator<T> getCreator() {
			return creator;
		}

		@Override
		protected void update(AnchorType type) {
			this.creator.getImageDto().anchorType = type;
		}

	}

	public static class PoiImageDto {
		
		public PoiRect padding = new PoiRect(0,0,0,0);
		
		public AnchorType anchorType = AnchorType.MOVE_AND_RESIZE;
	}
}
