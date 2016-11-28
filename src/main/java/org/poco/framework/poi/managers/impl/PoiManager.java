/**
 *
 */
package org.poco.framework.poi.managers.impl;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.poco.framework.poi.converter.IValueConverter;
import org.poco.framework.poi.converter.impl.ValueConverter;
import org.poco.framework.poi.creator.comment.ICommentCreator;
import org.poco.framework.poi.creator.image.IImageCreator;
import org.poco.framework.poi.dto.PoiPosition;
import org.poco.framework.poi.dto.PoiRect;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.exception.PoiException;
import org.poco.framework.poi.factory.CommentCreatorFactory;
import org.poco.framework.poi.factory.ImageCreatorFactory;
import org.poco.framework.poi.factory.PosFactory;
import org.poco.framework.poi.factory.StyleFactory;
import org.poco.framework.poi.factory.StyleManagerFactory;
import org.poco.framework.poi.managers.IPoiManager;
import org.poco.framework.poi.managers.IReadWriter;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.utils.PropertyUtil;
import org.poco.framework.poi.utils.PropertyUtil.IProperty;
import org.poco.framework.poi.utils.WorkbookUtil;

/**
 * POI操作用流れるようなインターフェース
 * @author yu-ki
 *
 */
public class PoiManager implements IPoiManager {

	private static Map<File,IPoiBook> cache = new HashMap<File, IPoiBook>();
	private static Map<Workbook,IPoiBook> cacheBook = new HashMap<Workbook, IPoiBook>();

	public static synchronized IPoiBook getInstance(File f) throws Exception {
		return getInstance(f,false);
	}
	/**
	 * Fileをキーにインスタンスを取得
	 * @param f
	 * @return
	 * @throws Exception
	 */
	public static synchronized IPoiBook getInstance(File f,Boolean isLowMemory) throws Exception {
		if (f == null) {
			throw new Exception("File is null.");
		}

		IPoiBook instance = cache.get(f);

		if (instance == null) {
			instance = new PoiBook(getWorkbook(f,isLowMemory),f);
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
	public static synchronized IPoiBook getInstance(Workbook book) throws Exception {
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
	public static void clearInstance(Workbook book) {
		if (book == null) {
			return;
		}
		IPoiBook instance = cacheBook.get(book);
		cacheBook.remove(instance);
	}

	/**
	 * POIのWorkBookを取得します。
	 * @param f ファイルが存在しない場合は、新規作成（ファイル拡張子に依存して作成するタイプを変更する）<BR/>
	 *          nullが指定された場合は、旧バージョン（XLS）
	 * @return
	 */
	public static Workbook getWorkbook(File f, Boolean isLowMemory) throws Exception {
		Workbook result = null;
		if (f == null) {
			//新規生成
			result = WorkbookUtil.createWorkBook("",isLowMemory);
		}
		//新規生成
		if (!f.exists()) {
			result = WorkbookUtil.createWorkBook(f.getName(),isLowMemory);
		}
		else {
			result = WorkbookUtil.readWorkbook(f,isLowMemory);
		}
		return result;
	}
	/**
	 * ブック
	 * @author yu-ki
	 */
	public static class PoiBook implements IPoiBook {

		private Workbook _book = null;
		private File _file = null;
		private IReadWriter rw;
		private Map<String,IPoiSheet> map = new HashMap<String, IPoiSheet>();

		/**
		 * コンストラクタ
		 * @param book
		 */
		public PoiBook(Workbook book, File file) {
			this._book = book;
			this._file = file;
		}

		public boolean saveBook() {
			return saveBook(_file);
		}

		public boolean saveBook(File f) {
			boolean result = true;
			try {
				boolean isXlsx = WorkbookUtil.isXlsx(_book);
				boolean exists = f.exists();
				WorkbookUtil.saveWorkBook(_book,f);
				//Poi　Bug: xlsxで、新規作成時に保存した場合、workBookがおかしくなるためリロードしておく
				if (isXlsx && !exists) {
					//スタイル破棄
					StyleFactory.removeStyle(this);
					//キャッシュ破棄
					map.clear();
					Boolean isLowMemory = (_book instanceof SXSSFWorkbook);
					//リロード
					_book = WorkbookUtil.readWorkbook(f,isLowMemory);
				}
				_file = f;
			}
			catch(Exception e) {
				result = false;
			}
			return result;
		}

		public Workbook getOrgWorkBook() {
			return _book;
		}

		public IPoiSheet cloneSheet(String fromName, String toName) {

			int index = this.getOrgWorkBook().getSheetIndex(fromName);
			//fromが存在しており、toが存在していないこと
			if (index != -1 && !WorkbookUtil.exists(this.getOrgWorkBook(), toName)) {

				//クローン作成
				Sheet sht = this.getOrgWorkBook().cloneSheet(index);
				//リネーム
				WorkbookUtil.rename(sht, toName);
			}
			return sheet(toName);
		}


		public IPoiSheet sheet(String name) {
			IPoiSheet psheet = map.get(name);

			if (psheet == null) {
				Sheet sht = this.getOrgWorkBook().getSheet(name);
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
		/**
		 * xmlによるCellロケーションに対して、書き込み、読み込みを行う
		 */
		public IReadWriter getReadWriter() {
			if (rw == null) {
				rw = new ReadWriter(this);
			}
			return rw;
		}
	}

	/**
	 * シート
	 * @author yu-ki
	 *
	 */
	public static class PoiSheet implements IPoiSheet {

		private Sheet _sh = null;
		private Map<PoiPosition,IPoiCell> map = new HashMap<PoiPosition, IPoiCell>();
		private PoiBook _parent;

		/**
		 * コンストラクタ
		 * @param book
		 * @param sh
		 */
		public PoiSheet(Sheet sh, PoiBook parent) {
			this._sh = sh;
			this._parent = parent;
		}


		public Sheet getOrgSheet() {
			return this._sh;
		}

		public IPoiBook getBook() {
			return _parent;
		}

		public IPoiCell cell(int x, int y) {
			PoiPosition pos = PosFactory.getInstance(x, y);
			IPoiCell cell = map.get(pos);
			if (cell == null) {
				cell = new PoiCell(pos, this);
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
			PoiRect rect = WorkbookUtil.analyseRangeAddress(address,this);
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
			PoiRect rect = WorkbookUtil.analyseRangeAddress(address,this);
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


		public Workbook getOrgWorkBook() {
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
			PoiRect pos = WorkbookUtil.analyseSelectRow(address,this);
			if (pos == null) {
				throw new ArrayIndexOutOfBoundsException("error[" + address + "]");
			}
			return insertRows(pos.top, pos.bottom - pos.top + 1);
		}


		public IReadWriter getReadWriter() {
			return _parent.getReadWriter();
		}
	}

	/**
	 * 範囲
	 * @author yu-ki
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
	 * @author yu-ki
	 *
	 */
	public static class PoiCell extends SheetChildBase implements IPoiCell {
		private static final int row_multiple = 20;
		private static final int column_multiple = 276;

		private Row _hRow = null;
		private Cell _hCell = null;
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
		public PoiCell(PoiPosition pos,IPoiSheet parent) {
			this._pos = pos;
			this._parent = parent;
		}

		private ValueConverter getValueManager() {
			if (_val == null) {
				_val = new ValueConverter(this);
			}
			return _val;
		}


		public Cell getOrgCell() {
			if (_hCell == null) {
				_hCell = WorkbookUtil.getCell(getOrgRow(), x());
			}
			return _hCell;
		}


		public Row getOrgRow() {
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

		/**
		 * 指定されたセルの座標からオブジェクトを縦横に展開する
		 */
		public IObjectWriter writeObject(Object dto) throws PoiException {
			List<Object> list = new ArrayList<Object>();
			list.add(dto);
			return writeObject(list);
		}
		/**
		 * 指定されたセルの座標からオブジェクトを縦横に展開する
		 */
		public IObjectWriter writeObject(List<Object> list) throws PoiException {

			return new PoiObjectWriter(this,list);
		}
	}

	public static class PoiObjectWriter implements IObjectWriter {

		private IPoiCell _cell;
		private List<Object> _headerList;
		private List<Object> _list;
		private List<String> _propertiesList;

		private List<String> getOrderList() throws PoiException {

			List<String> result = _propertiesList;
			if (result != null) {
				return result;
			}
			result = new ArrayList<String>();
			Object dto = _headerList != null ? _headerList.size() > 0 ? _headerList.get(0) : null : null;
			if (dto == null) {
				if (_list.size() > 0) {
					dto = _list.get(0);
				}
			}
			if (dto == null) throw new PoiException("write data not found.");

			if (dto instanceof Map<?, ?>) {
				Map<?,?> map = (Map<?,?>)dto;
				for (Object value : map.keySet()) {
					result.add(value.toString());
				}
				return result;
			}

			List<IProperty> props = PropertyUtil.getPropertiesAll(dto);
			for (IProperty prop : props) {
				result.add(prop.getName());
			}
			return result;
		}

		/**
		 * コンストラクタ
		 * @param cell
		 * @param dto
		 */
		public PoiObjectWriter(IPoiCell cell,List<Object> list) {
			_cell = cell;
			_list = list;
		}

		/**
		 * 指定条件に従い、データをPoiBookに書き込む
		 */
		public IPoiCell write() throws PoiException {
			if (_list.size() == 0 ) return _cell;

			List<String> orderList = getOrderList();

			Integer posX = _cell.x();
			Integer posY = _cell.y();

			Integer x,y = posY;

			if (_headerList != null) {
				for (Object header : _headerList) {
					x = posX;
					for (String key : orderList) {
						_cell.cell(x, y).setValue(getValue(header, key));
						x++;
					}
					y++;
				}
			}

			for (Object item : _list) {
				x = posX;
				for (String key : orderList) {
					_cell.cell(x, y).setValue(getValue(item, key));
					x++;
				}
				y++;
			}
			return _cell;
		}

		/**
		 * ヘッダを設定します
		 */
		public IObjectWriter setHeader(Object header) {
			if (header != null) {
				return setHeader(Arrays.asList(header));
			}
			return this;
		}
		/**
		 * ヘッダを設定します。複数行指定
		 */
		@Override
		public IObjectWriter setHeader(List<Object> headerList) {
			_headerList = headerList;
			return this;
		}
		/**
		 * 並び順、出力フィールドを設定する
		 */
		public IObjectWriter order(List<String> propertiesList)
				throws PoiException {
			_propertiesList = propertiesList;
			return this;
		}
		public IObjectWriter order(String[] propertiesArray)
				throws PoiException {
			List<String> order = new ArrayList<String>();
			for (String name : propertiesArray) {
				order.add(name);
			}
			return order(order);
		}

		private Object getValue(Object item, String propName) {
			Object result = null;
			if (item instanceof Map<?,?>) {
				Map<?,?> map = (Map<?,?>)item;
				result = map.get(propName);
			}
			else {
				result = PropertyUtil.getValue(item, propName);
			}
			//値がない場合は空白
			if (result == null) {
				return "";
			}
			return result;
		}
	}

	/**
	 * シートの子供ベース
	 * @author yu-ki
	 *
	 */
	public static class SheetChildBase implements IPoiSheet {
		protected IPoiSheet _parent;

		public IPoiBook getBook() {
			return _parent.getBook();
		}

		//--- sheetメソッド ----//

		public IPoiCell cell(int x, int y) {
			return _parent.cell(x,y);
		}


		public IPoiCell cell(String address) {
			return _parent.cell(address);
		}


		public Sheet getOrgSheet() {
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


		public Workbook getOrgWorkBook() {
			return _parent.getOrgWorkBook();
		}


		public IPoiSheet sheet(String name) {
			return _parent.sheet(name);
		}

		public IReadWriter getReadWriter() {
			return _parent.getReadWriter();
		}

	}

}
