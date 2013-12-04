/**
 * 
 */
package org.poco.framework.poi.managers.impl;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.poco.framework.poi.converter.IValueConverter;
import org.poco.framework.poi.converter.impl.ValueConverter;
import org.poco.framework.poi.creator.comment.ICommentCreator;
import org.poco.framework.poi.creator.image.IImageCreator;
import org.poco.framework.poi.dto.PoiPosition;
import org.poco.framework.poi.dto.PoiRect;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.factory.CommentCreatorFactory;
import org.poco.framework.poi.factory.ImageCreatorFactory;
import org.poco.framework.poi.factory.PosFactory;
import org.poco.framework.poi.factory.StyleFactory;
import org.poco.framework.poi.factory.StyleManagerFactory;
import org.poco.framework.poi.managers.IPoiManager;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.utils.WorkbookUtil;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * POI操作用流れるようなインターフェース
 * @author yu-ki106f
 *
 */
public class PoiManager implements IPoiManager {
	
	private static Map<File,IPoiBook> cache = new HashMap<File, IPoiBook>();
	private static Map<HSSFWorkbook,IPoiBook> cacheBook = new HashMap<HSSFWorkbook, IPoiBook>();
	
	/**
	 * Fileをキーにインスタンスを取得
	 * @param f
	 * @return
	 * @throws Exception
	 */
	public static synchronized IPoiBook getInstance(File f) throws Exception {
		if (f == null) {
			throw new Exception("File is null.");
		}
		
		IPoiBook instance = cache.get(f);

		if (instance == null) {
			instance = new PoiBook(getWorkbook(f),f);
			cache.put(f, instance);
		}
		return instance;
	}

	/**
	 * Fileをキーにインスタンスを取得
	 * @param f
	 * @return
	 * @throws Exception
	 */
	public static synchronized IPoiBook getInstance(HSSFWorkbook book) throws Exception {
		if (book == null) {
			throw new Exception("Workbook is null.");
		}
		
		IPoiBook instance = cacheBook.get(book);
		
		if (instance == null) {
			instance = new PoiBook(book, null);
			cacheBook.put(book, instance);
		}
		return instance;
	}

	/**
	 * 開放
	 */
	public static void destroy() {
		cache.clear();
		cacheBook.clear();
		PosFactory.clear();
		StyleFactory.destory();
		System.gc();
	}
	
	/**
	 * インスタンス破棄
	 * @param f
	 */
	public static void clearInstance(File f) {
		if (f == null) {
			return;
		}
		IPoiBook instance = cache.get(f);
		cache.remove(instance);
	}

	/**
	 * インスタンス破棄
	 * @param f
	 */
	public static void clearInstance(HSSFWorkbook book) {
		if (book == null) {
			return;
		}
		IPoiBook instance = cacheBook.get(book);
		cacheBook.remove(instance);
	}
	
	/**
	 * POIのWorkBookを取得します。
	 * @param f
	 * @return
	 */
	public static HSSFWorkbook getWorkbook(File f) throws Exception {
		HSSFWorkbook result = null;
		if (!f.exists()) {
			//新規生成
			result = new HSSFWorkbook();
		}
		else {
			result = WorkbookUtil.readWorkbook(f);
		}
		return result;
	}
	
	/**
	 * ブック
	 * @author yu-ki106f
	 */
	public static class PoiBook implements IPoiBook {

		private HSSFWorkbook _book = null;
		private File _file = null;
		private Map<String,IPoiSheet> map = new HashMap<String, IPoiSheet>();
		
		/**
		 * コンストラクタ
		 * @param book
		 */
		public PoiBook(HSSFWorkbook book, File file) {
			this._book = book;
			this._file = file;
		}

		public boolean saveBook() {
			return saveBook(_file);
		}

		public boolean saveBook(File f) {
			boolean result = true;
			try {
				WorkbookUtil.saveWorkBook(_book,f);
				_file = f;
			}
			catch(Exception e) {
				result = false;				
			}
			return result;
		}

		public HSSFWorkbook getOrgWorkBook() {
			return _book;
		}


		public IPoiSheet cloneSheet(String fromName, String toName) {
			
			int index = this.getOrgWorkBook().getSheetIndex(fromName);
			//fromが存在しており、toが存在していないこと
			if (index != -1 && !WorkbookUtil.exists(this.getOrgWorkBook(), toName)) {
				
				//クローン作成
				HSSFSheet sht = this.getOrgWorkBook().cloneSheet(index);
				//リネーム	
				WorkbookUtil.rename(sht, toName);
			}
			return sheet(toName);
		}


		public IPoiSheet sheet(String name) {
			IPoiSheet psheet = map.get(name);
			
			if (psheet == null) {
				HSSFSheet sht = this.getOrgWorkBook().getSheet(name);
				if (sht == null) {
					sht = this.getOrgWorkBook().createSheet();
					//リネーム
					WorkbookUtil.rename(sht, name);
				}
				psheet = new PoiSheet(sht, this);
				map.put(name, psheet);
			}
			
			return psheet;
		}
		
		/**
		 * シート名の変更
		 * @param name
		 * @param sheet
		 */
		public void renameSheet(String name, IPoiSheet sheet) {
			for (String key : map.keySet() ) {
				if (map.get(key).equals(sheet)) {
					map.remove(key);
					break;
				}
			}
			map.put(name, sheet);
		}

	}

	/**
	 * シート
	 * @author yu-ki106f
	 *
	 */
	public static class PoiSheet implements IPoiSheet {

		private HSSFSheet _sh = null;
		private Map<PoiPosition,IPoiCell> map = new HashMap<PoiPosition, IPoiCell>();
		private PoiBook _parent;
		/**
		 * コンストラクタ
		 * @param book
		 * @param sh
		 */
		public PoiSheet(HSSFSheet sh, PoiBook parent) {
			this._sh = sh;
			this._parent = parent;
		}
		

		public HSSFSheet getOrgSheet() {
			return this._sh;
		}


		public IPoiCell cell(int x, int y) {
			PoiPosition pos = PosFactory.getInstance(x, y);
			IPoiCell cell = map.get(pos);
			if (cell == null) {
				//HSSFRow hRow = WorkbookUtil.getRow(getOrgSheet(), pos.y);
				//HSSFCell hCell = WorkbookUtil.getCell(hRow, pos.x);
				cell = new PoiCell(/*hRow, hCell,*/ pos, this);
				map.put(pos, cell);
			}
			return cell;
		}


		public String sheetName() {
			return this.getOrgSheet().getSheetName();
		}


		public IPoiSheet setSheetName(String name) {
			//リネーム
			WorkbookUtil.rename(this.getOrgSheet(), name);
			//親の管理キャッシュもリネーム
			_parent.renameSheet(name, this);
			return this;
		}


		public IPoiCell cell(String address) {
			PoiPosition pos = WorkbookUtil.analyseAddress(address);
			return cell(pos.x, pos.y);
		}


		public IPoiRange range(String address) {
			PoiRect rect = WorkbookUtil.analyseRangeAddress(address);
			//範囲
			return range(rect);
		}
		

		public IPoiRange range(PoiRect rect) {
			return createRange(rect.left,rect.top,rect.right,rect.bottom, false);
		}
		

		public IPoiRange range(int fromX, int fromY, int toX, int toY) {
			return createRange(fromX,fromY,toX,toY, false);
		}
		
		/**
		 * レンジデータを作成
		 * @param fromX
		 * @param fromY
		 * @param toX
		 * @param toY
		 * @param merged
		 * @return
		 */
		private IPoiRange createRange(int fromX, int fromY, int toX, int toY, boolean merged) {
			List<IPoiCell> list = new ArrayList<IPoiCell>();
			for (int x = fromX; x <= toX; x++) {
				for (int y = fromY; y <= toY; y++) {
					list.add(cell(x, y));
				}
			}
			return new PoiRange(list.toArray(new IPoiCell[list.size()]), this, merged);
		}
		

		public IPoiSheet selected() {
			WorkbookUtil.setSelectAndActive(this.getOrgSheet());
			return this;
		}


		public IPoiRange merge(String address) {
			PoiRect rect = WorkbookUtil.analyseRangeAddress(address);
			//範囲
			return merge(rect);
		}
		

		public IPoiRange merge(PoiRect rect) {
			return merge(rect.left,rect.top,rect.right,rect.bottom);
		}


		public IPoiRange merge(int fromX, int fromY, int toX, int toY) {
			this.getOrgSheet().addMergedRegion(new CellRangeAddress(fromY,toY,fromX,toX));
			return createRange(fromX,fromY,toX,toY,true);
		}


		public boolean saveBook(File f) {
			return _parent.saveBook(f);
		}


		public boolean saveBook() {
			return _parent.saveBook();
		}


		public IPoiSheet cloneSheet(String fromName, String toName) {
			return _parent.cloneSheet(fromName, toName);
		}


		public HSSFWorkbook getOrgWorkBook() {
			return _parent.getOrgWorkBook();
		}


		public IPoiSheet sheet(String name) {
			return _parent.sheet(name);
		}


		public IPoiCell cell(PoiPosition pos) {
			return cell(pos.x,pos.y);
		}
		/**
		 * 指定座標から1行分、行の挿入を実施し、挿入した範囲を返す
		 * @param index Y座標
		 * @return 範囲
		 */

		public IPoiRange insertRows(int insertPos) {
			return insertRows(insertPos, 1);
		}
		
		/**
		 * 指定座標から指定行分、行の挿入を実施し、挿入した範囲を返す
		 * @param index Y座標
		 * @return 範囲
		 */

		public IPoiRange insertRows(int insertPos, int insertSize) {
			PoiPosition endPoint = WorkbookUtil.getEndPoint(this.getOrgSheet());
			this.getOrgSheet().shiftRows(insertPos, endPoint.y + insertSize, insertSize);
			//shiftのみしてレンジを指定する
			return range(0, insertPos, 255, insertPos + insertSize - 1);
		}
		
		/**
		 * 指定座標から指定行分、行の挿入を実施し、挿入した範囲を返す
		 * @param index Y座標
		 * @return 範囲
		 */

		public IPoiRange insertRows(String address) {
			PoiRect pos = WorkbookUtil.analyseSelectRow(address);
			if (pos == null) {
				throw new ArrayIndexOutOfBoundsException("error[" + address + "]");
			}
			return insertRows(pos.top, pos.bottom - pos.top + 1);
		}
		
	}

	/**
	 * 範囲
	 * @author yu-ki106f
	 */
	public static class PoiRange  extends SheetChildBase implements IPoiRange {
		private IPoiCell[] _cells;
		private boolean _merged = false;
		
		/**
		 * コンストラクタ
		 * @param cells
		 * @param parent
		 */
		public PoiRange(IPoiCell[] cells, IPoiSheet parent, boolean merged) {
			this._cells = cells;
			this._parent = parent;
			this._merged = merged;
		}
		

		public IPoiCell[] getCells() {
			return _cells;
		}


		public IStyleManager<IPoiRange> style() {
			return StyleManagerFactory.create(this);
		}

		public IStyleManager<IPoiRange> style(PoiStyleDto dto) {
			return StyleManagerFactory.create(this, dto);
		}


		public List<String> getStringValue() {
			List<String> list = new ArrayList<String>();
			for (IPoiCell cell : _cells) {
				list.add(cell.getStringValue());
			}
			return list;
		}


		public List<Object> getValue() {
			List<Object> list = new ArrayList<Object>();
			for (IPoiCell cell : _cells) {
				list.add(cell.getValue());
			}
			return list;
		}


		public <T> IPoiRange setValue(Object value, Class<T> clazz) {
			for (IPoiCell cell : _cells) {
				cell.setValue(value, clazz);
				//マージの場合は先頭のみ代入
				if (isMerged()) {
					break;
				}
			}
			return null;
		}
		

		public IPoiRange setValue(Object value) {
			for (IPoiCell cell : _cells) {
				cell.setValue(value);
				//マージの場合は先頭のみ代入
				if (isMerged()) {
					break;
				}
			}
			return this;
		}


		public Boolean isMerged() {
			return _merged;
		}
		

		public IPoiRange setColumnWidth(double value) {
			int lastX = -1;
			for (IPoiCell item : getCells()) {
				//x座標が変わったときのみ実施
				if (lastX != item.x()) {
					item.setColumnWidth(value);
					lastX = item.x();
				}
			}
			return this;
		}


		public IPoiRange setRowWidth(double value) {
			int lastY = -1;
			for (IPoiCell item : getCells()) {
				//y座標が変わったときのみ実施
				if (lastY != item.y()) {
					item.setRowWidth(value);
					lastY = item.y();
				}
			}
			return this;
		}

		public IPoiRange setColumnHidden(boolean hidden) {
			int lastX = -1;
			for (IPoiCell item : getCells()) {
				//x座標が変わったときのみ実施
				if (lastX != item.x()) {
					item.setColumnHidden(hidden);
					lastX = item.x();
				}
			}
			return this;
		}


		public IPoiRange setRownHidden(boolean hidden) {
			int lastY = -1;
			for (IPoiCell item : getCells()) {
				//y座標が変わったときのみ実施
				if (lastY != item.y()) {
					item.setRownHidden(hidden);
					lastY = item.y();
				}
			}
			return this;
		}


		public IPoiRange autoSizeColumn() {
			int lastX = -1;
			for (IPoiCell item : getCells()) {
				//x座標が変わったときのみ実施
				if (lastX != item.x()) {
					item.autoSizeColumn();
					lastX = item.x();
				}
			}
			return this;
		}


		public IPoiRange autoSizeColumn(boolean useMergedCells) {
			int lastX = -1;
			for (IPoiCell item : getCells()) {
				//x座標が変わったときのみ実施
				if (lastX != item.x()) {
					item.autoSizeColumn(useMergedCells);
					lastX = item.x();
				}
			}
			return this;
		}

		public IImageCreator<IPoiRange> setImage(File f) {
			return ImageCreatorFactory.create(this, f);
		}

		public IImageCreator<IPoiRange> setImage(byte[] bytes, String fileName) {
			return ImageCreatorFactory.create(this, bytes, fileName);
		}

		public ICommentCreator<IPoiRange> comment(String comment) {
			return CommentCreatorFactory.create(this, comment);
		}

	}
	
	/**
	 * セル
	 * @author yu-ki106f
	 *
	 */
	public static class PoiCell extends SheetChildBase implements IPoiCell {
		private static final int row_multiple = 20;
		private static final int column_multiple = 276;
		
		private HSSFRow _hRow = null;
		private HSSFCell _hCell = null;
		private PoiPosition _pos = null;
		private ValueConverter _val = null;
		
		private double _lastRowWidth = 12.75;
			
		public Integer x() {
			return	this._pos.x;
		}
		public Integer y() {
			return	this._pos.y;
		}

		/**
		 * コンストラクタ
		 * @param pos
		 * @param parent
		 */
		public PoiCell(/*HSSFRow row, HSSFCell cell,*/ PoiPosition pos,IPoiSheet parent) {
			//row と　cellは使用されたときに生成するように修正
			//this.hRow = row;
			//this.hCell = cell;
			this._pos = pos;
			this._parent = parent;
		}
		
		private ValueConverter getValueManager() {
			if (_val == null) {
				_val = new ValueConverter(this);
			}
			return _val;
		}
		

		public HSSFCell getOrgCell() {
			if (_hCell == null) {
				_hCell = WorkbookUtil.getCell(getOrgRow(), x());
			}
			return _hCell;
		}


		public HSSFRow getOrgRow() {
			if (_hRow == null) {
				_hRow = WorkbookUtil.getRow(getOrgSheet(), y());
			}
			return _hRow;
		}
		

		public IValueConverter value() {
			return getValueManager();
		}
		

		public String getStringValue() {
			return getValueManager().getString();
		}

		public Object getValue() {
			return getValueManager().getOriginalValue();
		}


		public IPoiCell setValue(Object value) {
			getValueManager().setValue(value);
			return this;
		}
		

		public <T> IPoiCell setValue(Object value, Class<T> clazz) {
			getValueManager().setValue(value,clazz);
			return this;
		}
		

		public IPoiCell setError(Byte value) {
			getValueManager().setError(value);
			return this;
		}
		

		public IPoiCell setFormura(String value) {
			getValueManager().setFormura(value);
			return this;
		}
		

		public IStyleManager<IPoiCell> style() {
			return StyleManagerFactory.create(this);
		}

		public IStyleManager<IPoiCell> style(PoiStyleDto dto) {
			return StyleManagerFactory.create(this, dto);
		}
		

		public IPoiCell setColumnHidden(boolean hidden) {
			_parent.getOrgSheet().setColumnHidden(x(), hidden);
			return this;
		}
		

		public IPoiCell setRownHidden(boolean hidden) {
			if (hidden) {
				if (this.getOrgRow().getHeight() != 0) {
					_lastRowWidth = this.getOrgRow().getHeight();
					_lastRowWidth /= row_multiple;
				}
				setRowWidth(0);
			}
			else {
				setRowWidth(_lastRowWidth);
			}
			return this;
		}


		public IPoiCell autoSizeColumn() {
			_parent.getOrgSheet().autoSizeColumn(x());
			return this;
		}
		

		public IPoiCell autoSizeColumn(boolean useMergedCells) {
			_parent.getOrgSheet().autoSizeColumn(x(), useMergedCells);
			return this;
		}


		public IPoiCell setColumnWidth(double point) {
			//276倍でそれなりに近い値にはなる。
			int value = Double.valueOf(point * column_multiple).intValue();
			_parent.getOrgSheet().setColumnWidth(x(), value);
			return this;
		}
		

		public IPoiCell setRowWidth(double point) {
			//20倍でサイズに変換する
			int value = Double.valueOf(point * row_multiple).intValue();
			if (Short.MAX_VALUE < value || Short.MIN_VALUE > value) {
				throw new java.lang.ClassCastException("short over parametor");
			}
			this.getOrgRow().setHeight((short)value);
			return this;
		}

		public IImageCreator<IPoiCell> setImage(File f) {
			return ImageCreatorFactory.create(this, f);
		}
		
		public IImageCreator<IPoiCell> setImage(byte[] bytes, String fileName) {
			return ImageCreatorFactory.create(this, bytes, fileName);
		}
		
		public ICommentCreator<IPoiCell> comment(String comment) {
			return CommentCreatorFactory.create(this, comment);
		}
	}

	/**
	 * シートの子供ベース
	 * @author yu-ki106f
	 *
	 */
	public static class SheetChildBase implements IPoiSheet {
		protected IPoiSheet _parent;
	

		//--- sheetメソッド ----//

		public IPoiCell cell(int x, int y) {
			return _parent.cell(x,y);
		}


		public IPoiCell cell(String address) {
			return _parent.cell(address);
		}


		public HSSFSheet getOrgSheet() {
			return _parent.getOrgSheet();
		}


		public IPoiRange merge(String address) {
			return _parent.merge(address);
		}


		public IPoiRange merge(int fromX, int fromY, int toX, int toY) {
			return _parent.merge(fromX, fromY, toX, toY);
		}


		public String sheetName() {
			return _parent.sheetName();
		}


		public IPoiRange range(String address) {
			return _parent.range(address);
		}


		public IPoiRange range(int fromX, int fromY, int toX, int toY) {
			return _parent.range(fromX, fromY, toX, toY);
		}


		public IPoiCell cell(PoiPosition pos) {
			return _parent.cell(pos);
		}


		public IPoiRange merge(PoiRect rect) {
			return _parent.merge(rect);
		}


		public IPoiRange range(PoiRect rect) {
			return _parent.range(rect);
		}
		

		public IPoiSheet selected() {
			return _parent.selected();
		}


		public IPoiSheet setSheetName(String name) {
			return _parent.setSheetName(name);
		}

		public IPoiRange insertRows(int insertPos) {
			return _parent.insertRows(insertPos);
		}

		public IPoiRange insertRows(int insertPos, int insertSize) {
			return _parent.insertRows(insertPos, insertSize);
		}

		public IPoiRange insertRows(String address) {
			return _parent.insertRows(address);
		}
		
		//---- book メソッド ----//

		public boolean saveBook(File f) {
			return _parent.saveBook(f);
		}


		public boolean saveBook() {
			return _parent.saveBook();
		}


		public IPoiSheet cloneSheet(String fromName, String toName) {
			return _parent.cloneSheet(fromName, toName);
		}


		public HSSFWorkbook getOrgWorkBook() {
			return _parent.getOrgWorkBook();
		}


		public IPoiSheet sheet(String name) {
			return _parent.sheet(name);
		}
	}

}
