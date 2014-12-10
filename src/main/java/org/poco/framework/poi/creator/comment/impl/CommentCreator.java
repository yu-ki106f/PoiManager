/**
 * 
 */
package org.poco.framework.poi.creator.comment.impl;

import org.poco.framework.poi.creator.comment.ICommentCreator;
import org.poco.framework.poi.dto.PoiRect;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiRange;
import org.poco.framework.poi.utils.WorkbookUtil;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;

/**
 * @author funahashi
 *
 */
public class CommentCreator<T> implements ICommentCreator<T> {

	private String comment = null;
	
	private String createUser = null;
	
	public T parent = null;
	
	private PoiRect padding = null;

	private PoiRect startDiff = null;
	
	public CommentCreator(String comment) {
		this.comment = comment;
		//初期値 all:0
		padding(0);
		//Cellポイント x軸 + 1, +3 , y軸 -1 , + 2
		startPosDiff(1, -1, 3, 2);
		
	}

	public ICommentCreator<T> startPosDiff(int xStart, int yStart, int xEnd, int yEnd) {
		this.startDiff = new PoiRect();
		this.startDiff.bottom = yEnd;
		this.startDiff.left = xStart;
		this.startDiff.right = xEnd;
		this.startDiff.top = yStart;
		return this;
	}

	public ICommentCreator<T> padding(int left, int top, int right, int bottom) {
		this.padding = new PoiRect();
		this.padding.bottom = bottom;
		this.padding.left = left;
		this.padding.right = right;
		this.padding.top = top;
		return this;
	}

	/**
	 * @see org.poco.framework.poi.creator.comment.ICommentCreator#padding(org.poco.framework.poi.dto.PoiRect)
	 */
	public ICommentCreator<T> padding(PoiRect rect) {
		this.padding = rect;
		return this;
	}

	/**
	 * @see org.poco.framework.poi.creator.comment.ICommentCreator#padding(int)
	 */
	public ICommentCreator<T> padding(int all) {
		return padding(all,all,all,all);
	}

	public ICommentCreator<T> comment(String value) {
		this.comment = value;
		return this;
	}

	public ICommentCreator<T> createUser(String value) {
		this.createUser = value;
		return this;
	}
	
	/**
	 * @see org.poco.framework.poi.creator.comment.ICommentCreator#write()
	 */
	public T write() {
		if (parent instanceof IPoiRange) {
			IPoiRange poiRange = (IPoiRange)parent;
			for (IPoiCell cell : poiRange.getCells()) {
				setComment(cell);
			}
		}
		else {
			setComment((IPoiCell)parent);
		}
		return parent;
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
	
	private void setComment(IPoiCell cell) {
		Drawing patriarch = null;
		PoiRect pos = getPosition(cell);
		
		try {
			//コメントカラムを取得し、存在しない場合は追加する
			Comment hComment = cell.getOrgCell().getCellComment();
			CreationHelper createHelper = getSheet().getWorkbook().getCreationHelper();
			if (hComment == null) {
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
				clientAnchor.setDx1(this.padding.left);
				clientAnchor.setDy1(this.padding.top);
	
				clientAnchor.setDx2(this.padding.right);
				clientAnchor.setDy2(this.padding.bottom);
				
				clientAnchor.setCol1(pos.left);
				clientAnchor.setRow1(pos.top);
				clientAnchor.setCol2(pos.right);
				clientAnchor.setRow2(pos.bottom);
	
				//コメントを生成する
				hComment = patriarch.createCellComment(clientAnchor);
				cell.getOrgCell().setCellComment(hComment);
			}
			//コメントを設定する
			hComment.setString(createHelper.createRichTextString(comment));
			//コメント作成者を設定する
			if (createUser != null) {
				hComment.setAuthor(createUser);
			}
		}
		catch(Exception e) {
			//TODO ERROR
			e.printStackTrace();
		}
	}
	
	private PoiRect getPosition(IPoiCell cell) {
		//Cellポイント x軸 + 1, +3 , y軸 -1 , + 2(デフォルト）
		PoiRect result = new PoiRect();

		//X軸開始
		result.left = checkPos(0, 255, cell.x() + this.startDiff.left);
		//X軸終了
		result.right= checkPos(0, 255, cell.x() + this.startDiff.right);
		//Y軸開始
		result.top = checkPos(0,65535,cell.y() + this.startDiff.top);
		//Y軸終了
		result.bottom = checkPos(0, 65535, cell.y() + this.startDiff.bottom);
		
		return result;
		
	}

	/**
	 * 範囲チェック
	 * @param scopeStart
	 * @param scopeEnd
	 * @param value
	 * @return
	 */
	private Integer checkPos(Integer scopeStart, Integer scopeEnd, Integer value) {
		if (value < scopeStart) {
			return scopeStart;
		}
		if (value > scopeEnd) {
			return scopeEnd;
		}
		return value;
	}
}
