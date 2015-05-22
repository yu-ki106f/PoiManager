package org.poco.framework.poi.managers.impl;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.poco.framework.poi.converter.IConverter;
import org.poco.framework.poi.dto.PoiPosition;
import org.poco.framework.poi.exception.PoiException;
import org.poco.framework.poi.managers.IPoiManager.IPoiBook;
import org.poco.framework.poi.managers.IPoiManager.IPoiCell;
import org.poco.framework.poi.managers.IPoiManager.IPoiSheet;
import org.poco.framework.poi.managers.IReadWriter;
import org.poco.framework.poi.managers.IStyleManager;
import org.poco.framework.poi.utils.ConvertUtil;
import org.poco.framework.poi.utils.PropertyUtil;
import org.poco.framework.poi.utils.WorkbookUtil;
import org.poco.framework.poi.utils.XMLUtils;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.xml.sax.SAXException;

public class ReadWriter implements IReadWriter {
	
	private IPoiBook book;
	
	private XMLBook templeteXml = null;
	
	//Listの先頭行のスタイルのコピーの有無
	private Boolean styleCopy = false;
	
	public static Integer MAX_ROWS = 65535;

	/**
	 * テンプレートのクローンを返す 
	 * @return
	 */
	public XMLBook getTempleteXMLClone() {
		if (templeteXml != null) {
			return templeteXml.clone();
		}
		return null;
	}
	public IReadWriter setMaxRows(Integer value) {
		MAX_ROWS = value;
		return this;
	}
	public IReadWriter setStyleCopy(boolean value) {
		styleCopy = value;
		return this;
	}
	public Integer getMaxRows() {
		return MAX_ROWS;
	}
	
	public ReadWriter(IPoiBook value) {
		book = value;
	}
	public IPoiSheet cloneSheet(String fromName, String toName) {
		return book.cloneSheet(fromName, toName);
	}
	public boolean saveBook(File f) {
		return book.saveBook(f);
	}
	public boolean saveBook() {
		return book.saveBook();
	}
	public IPoiSheet sheet(String name) {
		return book.sheet(name);
	}
	public Workbook getOrgWorkBook() {
		return book.getOrgWorkBook();
	}
	public boolean isXlsx() {
		return WorkbookUtil.isXlsx(getOrgWorkBook());
	}
	
	/**
	 * PoiBookに指定されたデータを書き込み
	 * @param data
	 * @param templete テンプレートデータを指定
	 * @return
	 * @throws PoiException
	 */
	public IPoiBook write(Object data, XMLBook templete) throws PoiException {
		if (templete == null) {
			throw new PoiException("It is not crowded read the XML file.");
		}
		try {
			IPoiSheet sheet;
			for (XMLSheet xsheet : templete.sheets) {
				//シート取得
				sheet = book.sheet(xsheet.name);
				//入出力設定文字削除
				clearTemplete(sheet, xsheet);
				//個別項目書き込み
				onceWrite(sheet, xsheet.onces, data);
				//リスト項目書き込み
				listWrite(sheet, xsheet.lists, data);
			}
		} catch (Exception e) {
			e.printStackTrace();
			throw new PoiException(e.getMessage());
		}
		return book;
	}
	
	/**
	 * PoiBookに指定されたデータを書き込み
	 * @param data
	 * @return
	 * @throws PoiException 
	 */
	public IPoiBook write(Object data) throws PoiException {
		return write(data,templeteXml);
	}

	/**
	 * PoiBookから設定ファイルルールにてデータ読み込み
	 * @param <T>
	 * @param baseClass
	 * @param templete テンプレートデータを指定
	 * @return
	 * @throws PoiException
	 */
	public <T> T read(Class<T> baseClass, XMLBook templete) throws PoiException {
		T result = null;
		if (templete == null) {
			throw new PoiException("It is not crowded read the XML file.");
		}
		try {
			result = baseClass.newInstance();
			IPoiSheet sheet;
			for (XMLSheet xsheet : templete.sheets) {
				//シート取得
				sheet = book.sheet(xsheet.name);
				//個別項目読み込み
				onceRead(sheet, xsheet.onces, result, templete.defaultPackage);
				//リスト項目読み込み
				listRead(sheet, xsheet.lists, result, templete.defaultPackage);
			}
		} catch (Exception e) {
			e.printStackTrace();
			throw new PoiException(e.getMessage());
		}
		return result;
	}
	/**
	 * PoiBookから設定ファイルルールにてデータ読み込み
	 * @param <T>
	 * @param baseClass
	 * @return
	 */
	public <T> T read(Class<T> baseClass) throws PoiException {
		return read(baseClass,templeteXml);
	}
	
	/**
	 * 個別データをオブジェクトに格納する
	 * @param sheet　シート
	 * @param onces　個別入出力指定リスト
	 * @param data　オブジェクト
	 * @throws Exception
	 */
	private void onceRead(IPoiSheet sheet, List<XMLDetail> onces, Object data, String defaultPackage) throws Exception {
		Object dto;
		for (XMLDetail xdetail : onces) {
			dto = PropertyUtil.getValue(data, xdetail.name);
			if (dto == null) {
				dto = getDtoClass(xdetail.className,defaultPackage).newInstance();
				PropertyUtil.setValue(data, xdetail.name, dto);
			}
			//詳細書き込み
			for (XMLCell xcell : xdetail.cells) {
				cellRead(sheet, xcell, dto, null);
			}
		}
	}
	/**
	 * リストデータをデータが存在するまでオブジェクトリストに格納する
	 * @param sheet　シート
	 * @param lists リスト入出力指定リスト
	 * @param data　オブジェクト
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	private void listRead(IPoiSheet sheet, List<XMLDetail> lists, Object data, String defaultPackage) throws Exception {
		List<Object> list;
		Object dto;
		Integer yPos = 0;
		for (XMLDetail xdetail : lists) {
			list = (List<Object>)PropertyUtil.getValue(data, xdetail.name);
			if (list == null) {
				list = new ArrayList<Object>();
				PropertyUtil.setValue(data, xdetail.name, list);
			}
			//Rowが存在しなくなるまでループ
			while(isExists(sheet, xdetail, yPos)) {
				dto = getDtoClass(xdetail.className,defaultPackage).newInstance();
				//詳細書き込み

				for (XMLCell xcell : xdetail.cells) {
					cellRead(sheet, xcell, dto, new PoiPosition(0,yPos));
				}
				list.add(dto);
				yPos++;
			}
		}
	}
	/**
	 * 指定行のデータが存在するかを確認する
	 * @param sheet
	 * @param xdetail
	 * @param yPos
	 * @return
	 */
	private boolean isExists(IPoiSheet sheet,XMLDetail xdetail,Integer yPos) {
		for (Integer y : xdetail.getYPosition()) {
			if (sheet.getOrgSheet().getRow(y+yPos) == null) {
				return false;
			}
		}
		return true;
	}
	/**
	 * セルの内容を読み出し格納する
	 * @param sheet
	 * @param xcell
	 * @param value
	 * @param diff
	 * @throws Exception
	 */
	private void cellRead(IPoiSheet sheet, XMLCell xcell, Object value, PoiPosition diff) throws Exception {
		//セル取得
		IPoiCell cell = sheet.cell(mergePosion(xcell.point, diff));
		//コンバーター取得
		IConverter conv = ConvertUtil.getConverter(cell.value().getOriginalValue());
		//Dtoのタイプ取得
		Class<?> type = PropertyUtil.getType(value, xcell.name);
		//Dtoの型に変換
		Object data = ConvertUtil.convert(conv,type);
		//値を設定
		PropertyUtil.setValue(value, xcell.name, data, type);
	}
	/**
	 * 個別項目を書き込む
	 * @param sheet
	 * @param onces
	 * @param data
	 */
	private void onceWrite(IPoiSheet sheet, List<XMLDetail> onces, Object data) {
		Object param = null;
		for (XMLDetail xdetail : onces) {
			param = PropertyUtil.getValue(data, xdetail.name);
			if (param == null) continue;
			//詳細書き込み
			for (XMLCell xcell : xdetail.cells) {
				//文字
				Object value = PropertyUtil.getValue(param, xcell.name);
				if (value == null) continue;
				cellWrite(sheet, xcell, value, null);
			}
		}
	}
	
	/**
	 * リスト項目書き込み
	 * @param sheet
	 * @param lists
	 * @param data
	 */
	private void listWrite(IPoiSheet sheet, List<XMLDetail> lists, Object data) {
		List<?> list = null;
		Integer counter = 0;
		Integer pos = 0;
		Object item = null;
		Boolean isXlsx = isXlsx();
		List<IPoiSheet> sheets;
		IPoiSheet tmpSheet;
		IPoiCell tmpCell;
		for (XMLDetail xdetail : lists) {
			list = (List<?>)PropertyUtil.getValue(data, xdetail.name);
			if (list == null) continue;
			//シート複製
			sheets = createCloneSheet(sheet, list.size());
			counter = 0;
			Map<XMLCell,CellStyle> styles = new HashMap<XMLCell, CellStyle>();
			//件数分ループ
			for (Object param : list) {
				pos = isXlsx ? counter : counter % MAX_ROWS;
				tmpSheet = getTargetSheet(sheets, counter);
				//詳細書き込み
				for (XMLCell xcell : xdetail.cells) {
					item = PropertyUtil.getValue(param, xcell.name);
					//内容書き込み
					tmpCell = cellWrite(tmpSheet, xcell, item, new PoiPosition(0,pos));
					//スタイル適用
					if (styles.containsKey(xcell)) {
						((StyleManager<IPoiCell>)tmpCell.style()).forceUpdate(styles.get(xcell));
					}
					//2行目以降に反映されるようにする
					if (styleCopy) {	//コピー指定があるときのみ生成
						if (!styles.containsKey(xcell) ) {
							styles.put(xcell, tmpSheet.cell(xcell.point).getOrgCell().getCellStyle());
						}
					}
				}
				counter++;
			}
		}
	}
	/**
	 * 各セル項目書き込み
	 * @param sheet
	 * @param xcell
	 * @param value
	 * @param diff
	 */
	private IPoiCell cellWrite(IPoiSheet sheet, XMLCell xcell, Object value, PoiPosition diff) {
		IPoiCell cell = sheet.cell(mergePosion(xcell.point, diff));
		//色
		updateForegroundColor(cell.style(), xcell.color);
		//セルタイプ
		Class<?> clazz = getWriteClass(xcell.type);				
		if (clazz == null) {
			cell.setValue(value);
		}
		else {
			try {
				cell.setValue(value, clazz);
			}
			catch(Exception e) {
				cell.setValue(value);
			}
		}
		return cell;
	}
	/**
	 * positionマージ
	 * @param pos1
	 * @param pos2
	 * @return
	 */
	private PoiPosition mergePosion(PoiPosition pos1, PoiPosition pos2) {
		if (pos2 == null) {
			return pos1;
		}
		return new PoiPosition(pos1.x + pos2.x ,pos1.y + pos2.y);
	}
	/**
	 * 書き込みクラス取得
	 * @param type
	 * @return
	 */
	private Class<?> getWriteClass(String type){
		try {

			if (isEmpty(type) || "TEMPLATE".equals(type)) return null;
			//イリーガルは無視
			return CellType.valueOf(type).getClazz();
		}
		catch(Exception e){}
		return null;
	}
	
	/**
	 * 背景色を更新する
	 * @param style
	 * @param name
	 */
	private void updateForegroundColor(IStyleManager<?> style,String name) {
		try {
			if (isEmpty(name) || "TEMPLATE".equals(name)) return;
			//イリーガルは無視
			style.fill().foregroundColor().name(name).fill().pattern().SOLID_FOREGROUND().update();
		}
		catch(Exception e){}
	}
	/**
	 * 設定文字をクリアする
	 * @param sheet
	 * @param xsheet
	 */
	private void clearTemplete(IPoiSheet sheet, XMLSheet xsheet) {
		List<XMLDetail> list = cloneList(xsheet.onces);
		list.addAll(xsheet.lists);
		for (XMLDetail xdetail : list) {
			for (XMLCell xcell : xdetail.cells) {
				//クリア
				sheet.cell(xcell.point).setValue("");
			}
		}
	}
	/**
	 * 対象シートを取得する<BR/>
	 * XLSXの場合は1シートに全て
	 * @param sheets
	 * @param counter
	 * @return
	 */
	private IPoiSheet getTargetSheet(List<IPoiSheet> sheets, Integer counter) {
		if (isXlsx()) {
			return sheets.get(0);
		}
		//0スタートのため
		int index = (counter) / MAX_ROWS;
		return sheets.get(index);
	}
	/**
	 * 出力件数に依存して、複製シートを生成する（XLS専用）
	 * @param sheet
	 * @param length
	 * @return
	 */
	private List<IPoiSheet> createCloneSheet(IPoiSheet sheet,Integer length) {
		List<IPoiSheet> sheets = new ArrayList<IPoiSheet>();
		sheets.add(sheet);
		if (!isXlsx()) {
			Integer sheetCount = 1;
			if (length > MAX_ROWS) {
				sheetCount = Double.valueOf(Math.ceil( length.doubleValue() / MAX_ROWS )).intValue();
			}
			//clone作成
			for (Integer val = 1; val < sheetCount; val++) {
				sheets.add(sheet.cloneSheet(sheet.sheetName(), sheet.sheetName()+"_"+val.toString()));
			}
		}
		return sheets;
	}
	/**
	 * リストのクローンを生成する
	 * @param <T>
	 * @param list
	 * @return
	 */
	private <T> List<T> cloneList(List<T> list) {
		List<T> result = new ArrayList<T>();
		for (T item : list) {
			result.add(item);
		}
		return result;
	}
	
	private Class<?> getDtoClass(String className, String defaultPackage) throws PoiException{
		Class<?> result = _getClass(className,defaultPackage);
		if (result == null) {
			throw new PoiException("class not found. ["+className+"]");
		}
		return result;
	}
	private Class<?> _getClass(String className,String defaultPackage) {
		try {
			return Class.forName(className);
		}
		catch(Exception e) {}
		try {
			return Class.forName(defaultPackage.concat(".").concat(className));
		}
		catch(Exception e) {}
		return null;
	}
	//----------------------------XML読込--------------------------
	/**
	 * 例外が発生しなければ、読込み成功
	 */
	public IReadWriter loadFromXml(String fileName) throws PoiException {
		InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream(fileName);
		try {
			templeteXml = loadXml(inputStream);
		}
		catch(Exception e){
			throw new PoiException(e.getMessage());
		}
		finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {}
			}
		}
		return this;
	}
	private XMLBook loadXml(InputStream inputStream) throws IOException, SAXException, ParserConfigurationException {
		try{
			// Javax
			DocumentBuilder db = DocumentBuilderFactory.newInstance().newDocumentBuilder();
			Document doc = db.parse(inputStream);
			// ノードの取得
			Node root = doc.getDocumentElement();
			// book
			return createXMLBook(root);
		} catch (IOException e) {
			// ファイル読込みエラー
			throw e;
		} catch (SAXException e) {
			// XMLの形式が不正
			throw e;
		} catch (ParserConfigurationException e) {
			// XMLの形式が不正
			throw e;
		}
	}
	/**
	 * 指定されたNodeからXML解析を行い、結果としてXMLBookを生成して返す。
	 * @param root
	 * @return
	 */
	private XMLBook createXMLBook(Node root) {
		XMLBook result = new XMLBook();
		
		result.name = XMLUtils.getAttributeValue(root, "name");
		// デフォルトパッケージ名を取得
		result.defaultPackage = XMLUtils.getAttributeValue(root, "defaultPackage");
		// テンプレートファイルPathを取得
		result.path = XMLUtils.getAttributeValue(root, "path");
		// テンプレートファイルKeyを取得
		result.key = XMLUtils.getAttributeValue(root, "key");
		//シート一覧取得
		result.sheets = createXMLSheetList(root);
		
		return result;
	}
	private List<XMLSheet> createXMLSheetList(Node root) {
		List<XMLSheet> result = new ArrayList<XMLSheet>();
		
		for(Node sheet: XMLUtils.getChildNodeList(root)) {
			//シート追加
			result.add(createXMLSheet(sheet));
		}
		return result;
	}
	private XMLSheet createXMLSheet(Node sheet) {
		XMLSheet result = new XMLSheet();
		// シート名を取得
		result.name = XMLUtils.getAttributeValue(sheet, "name");
		result.lists = new ArrayList<XMLDetail>();
		result.onces = new ArrayList<XMLDetail>();
		
		// sheetの子供を取得
		for(Node detail : XMLUtils.getChildNodeList(sheet)) {
			// 可変
			if(detail.getNodeName().equals("list")){
				result.lists.add(createXMLDetail(detail));
			}
			else if(detail.getNodeName().equals("once")){
				result.onces.add(createXMLDetail(detail));
			}
		}
		return result;
	}
	private XMLDetail createXMLDetail(Node detail) {
		XMLDetail result = new XMLDetail();
		result.className = XMLUtils.getAttributeValue(detail, "className");
		result.name = XMLUtils.getAttributeValue(detail, "name");
		result.cells = createXMLCells(detail);
		return result;
	}
	private List<XMLCell> createXMLCells(Node detail) { 
		List<XMLCell> result = new ArrayList<XMLCell>();
		XMLCell cell;
		Node point;
		//cellたちを取得
		for(Node node:  XMLUtils.getChildNodeList(detail)) {
			cell = new XMLCell();
			cell.point = new PoiPosition();
			// 座標を取得
			point = XMLUtils.getChildNode(node, "point");
			cell.point.x = (Integer.parseInt(XMLUtils.getChildNodeText(point, "x")));
			cell.point.y = (Integer.parseInt(XMLUtils.getChildNodeText(point, "y")));
			cell.name = XMLUtils.getChildNodeText(node, "name");
			cell.type = XMLUtils.getChildNodeText(node, "type");
			cell.color = XMLUtils.getChildNodeText(node, "color");
			result.add(cell);
		}
		return result;
	}

	//----------------------------関連DTO--------------------------
	/**
	 * Excel入出力設定ファイル情報 Book
	 * @author neutral shibata
	 */
	public class XMLBook {
		/**
		 * ブック名(PoiManager未使用）
		 */
		public String name;
		/**
		 * デフォルトパッケージ名
		 */
		public String defaultPackage;
		/**
		 * テンプレートのPath(PoiManager未使用）
		 */
		public String path;
		/**
		 * テンプレートのKey(PoiManager未使用）
		 */
		public String key;
		/**
		 * Excel入出力設定ファイル情報 Sheet
		 */
		public List<XMLSheet> sheets;
		
		/**
		 * クローンメソッド
		 */
		public XMLBook clone() {
			XMLBook result = new XMLBook();
			result.name = this.name;
			result.defaultPackage = this.defaultPackage;
			result.path = this.path;
			result.key = this.key;
			if (this.sheets != null) {
				result.sheets = new ArrayList<XMLSheet>();
				for (XMLSheet sheet : this.sheets) {
					result.sheets.add(sheet.clone());
				}
			}
			return result;
		}
	}
	
	/**
	 * Excel入出力設定ファイル情報 Sheet
	 * @author neutral shibata
	 */
	public class XMLSheet {
		/**
		 * シート名
		 */
		public String name;
		/**
		 * Excel入出力設定ファイル情報 固定
		 */
		public List<XMLDetail> onces;
		/**
		 * Excel入出力設定ファイル情報 リスト
		 */
		public List<XMLDetail> lists;
		/**
		 * クローンメソッド
		 */
		public XMLSheet clone() {
			XMLSheet result = new XMLSheet();
			result.name = this.name;
			if (this.onces != null) {
				result.onces = new ArrayList<XMLDetail>();
				for (XMLDetail detail : this.onces) {
					result.onces.add(detail.clone());
				}
			}
			if (this.lists != null) {
				result.lists = new ArrayList<XMLDetail>();
				for (XMLDetail detail : this.lists) {
					result.lists.add(detail.clone());
				}
			}
			return result;
		}
	}
	
	/**
	 * Excel入出力設定ファイル情報 Detail
	 * @author neutral shibata
	 */
	public class XMLDetail {
		/**
		 * メンバ名称
		 */
		public String name;
		/**
		 * クラス名称
		 */
		public String className;
		/**
		 * Excel入出力設定ファイル情報 Cell
		 */
		public List<XMLCell> cells;
		
		private List<Integer> _yPos = new ArrayList<Integer>();
		
		public void clearYPotion() {
			_yPos = new ArrayList<Integer>();
		}
		
		public List<Integer> getYPosition() {
			if (_yPos == null || _yPos.size() == 0) {
				_yPos = new ArrayList<Integer>();
				for (XMLCell cell : cells) {
					if (!_yPos.contains(cell.point.y) ) {
						_yPos.add(cell.point.y);
					}
				}
			}
			return _yPos;
		}
		
		/**
		 * クローンメソッド
		 */
		public XMLDetail clone() {
			XMLDetail result = new XMLDetail();
			result.name = this.name;
			result.className = this.className;
			if (this.cells != null) {
				result.cells = new ArrayList<XMLCell>();
				for (XMLCell cell : this.cells) {
					result.cells.add(cell.clone());
				}
			}
			return result;
		}
	}

	/**
	 * Excel入出力設定ファイル情報 Cell
	 * @author neutral shibata
	 */
	public class XMLCell {
		/**
		 * 座標
		 */
		public PoiPosition point;
		/**
		 * メンバ名称
		 */
		public String name;
		/**
		 * セルのタイプ
		 */
		public String type;
		/**
		 * セルの色
		 */
		public String color;
		/**
		 * クローンメソッド
		 */
		public XMLCell clone() {
			XMLCell result = new XMLCell();
			result.color = this.color;
			result.name = this.name;
			result.type = this.type;
			result.point = new PoiPosition(this.point.x, this.point.y);
			return result;
		}
	}
	
	enum CellType {
		STRING(String.class),
		BLANK(String.class),
		BOOLEAN(Boolean.class),
		ERROR(String.class),
		NUMERIC(Double.class),
		FORMULA(String.class),
		DATE_TIME(Date.class),
		DATE(Date.class);
		private Class<?> clazz;
		CellType(Class<?> value){
			clazz = value;
		}
		public Class<?> getClazz(){
			return clazz;
		}
	}
	
	public boolean isEmpty(Object value) {
		if (value == null) return true;
		if (value.toString().equals("")) return true;
		return false;
	}
}
