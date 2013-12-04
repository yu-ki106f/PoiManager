package org.poco.framework.poi.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.poco.framework.poi.dto.PoiPosition;
import org.poco.framework.poi.dto.PoiRect;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.IOUtils;

public class WorkbookUtil {
	private static final String TYPE_JPG = "jpg";
	private static final String TYPE_JPEG = "jpeg";
	private static final String TYPE_PNG = "png";
	private static final String TYPE_DIB = "bmp";
	
	/**
	 * Excelファイル(直接Path指定)を取得する
	 * @param String ファイル名
	 * @return HSSFWorkbook 作成したブック
	 * @throws FljException 
	 */
	public static HSSFWorkbook readWorkbook(File f) throws Exception {
		
		FileInputStream is = null; 
		POIFSFileSystem psys = null;
		HSSFWorkbook result = null;
		try {
			
			// テンプレートを取得
			is = new FileInputStream(f);
			
			psys = new POIFSFileSystem(is);
			
			result = new HSSFWorkbook(psys);
			
		} catch (FileNotFoundException e) {
			throw e;
		} catch (IOException e) {
			throw e;
		} finally {
			try {
				if(is != null) {
					is.close();
				}
			}catch (IOException e){}
		}
		return result;
	}

	/**
	 * ファイルにExcel情報を出力する
	 * @param filePathName
	 * @throws FljException
	 */
	public static void saveWorkBook(HSSFWorkbook wb, File f) throws Exception {
		
		if (f == null ) {
			throw new Exception("保存先の設定が異常です。");
		}
		if (f.isDirectory()) {
			throw new Exception("保存先がディレクトリです。");
		}
		if (wb == null) {
			throw new Exception("保存内容がありません。");
		}
		
		try{
			FileOutputStream fos = null;
			try {
				// ファイルの書き出し処理
				fos = new FileOutputStream(f);
				
				// ワークブック書き込み
				wb.write(fos);
				
			} catch (FileNotFoundException e) {
				// 書き込み対象にアクセスできない
				throw e;
			}finally {
				
				// ファイルクローズ
				if (fos != null) {
					fos.close();
				}
			}
		} catch (IOException e) {
			// 書き込み対象にアクセスできない
			throw e;
		}
	}

	/**
	 * ExcelのRowを取得する。
	 * （行が無い場合新たに作成します）
	 * @param HSSFSheet POIのExcelSheetオブジェクト
	 * @param int 行番号
	 */
	public static HSSFRow getRow(HSSFSheet sheet, int i) {
		HSSFRow row = sheet.getRow(i);
		if (row == null) {
			row = sheet.createRow(i);
		}
		return row;
	}
	
	/**
	 * ExcelのCellを取得する。
	 * （セルが無い場合新たに作成します）
	 * @param HSSFRow POIのExcelRowオブジェクト
	 * @param int 列番号
	 */
	public static HSSFCell getCell(HSSFRow row, int i) {
		HSSFCell cell = row.getCell(i);
		if (cell == null) {
			cell = row.createCell(i);
		}
		return cell;
	}
	
	public static boolean rename(HSSFSheet srcSheet, String name) {
		if (srcSheet == null) return false;
		
		HSSFWorkbook book = srcSheet.getWorkbook();
		int index = book.getSheetIndex(srcSheet);
		//対象シートが存在する 且つ 変更後の名前のシートがないこと
		if (index != -1 && !exists(book, name)) {
			//リネーム
			srcSheet.getWorkbook().setSheetName(index, name);
			return true;
		}
		return false;
	}
	
	public static PoiRect analyseRangeAddress(String value) {

		if (value == null) return null;
		
		String[] values = value.split(":");
		
		if (values.length == 0) {
			throw new ArrayIndexOutOfBoundsException("error[" + value + "]");
		}
		else if (values.length == 1) {
			return new PoiRect(analyseAddress(values[0]));
		}
		else if (values.length == 2){
			//A:A or 1:1 (列、行指定の場合）
			PoiRect rect = analyseSelectColumnOrRow(value);
			if (rect != null) {
				return rect;
			}
			//通常パターン
			return new PoiRect(analyseAddress(values[0]), analyseAddress(values[1]) );
		}
		throw new ArrayIndexOutOfBoundsException("error[" + value + "]");
	}

	public static PoiRect analyseSelectColumn(String value) {
		if (value == null) return null;
		
		String tmp = value.replaceAll("\\$", "").toUpperCase();
		String columStr = tmp.replaceAll("(^[A-Z]+):([A-Z]+)$", "$1#$2");
		String values[];
		Integer pos1 = 0;
		Integer pos2 = 0;
		//列指定
		if (!tmp.equals(columStr)){
			values = columStr.split("#");
			if (values.length == 2 ) {
				pos1 = alpabetToNumber(values[0]);
				pos2 = alpabetToNumber(values[1]);
				return new PoiRect(0,pos1,pos2,65535);
			}
		}
		return null;
	}
	
	public static PoiRect analyseSelectRow(String value) {
		if (value == null) return null;
		
		String tmp = value.replaceAll("\\$", "").toUpperCase();
		String rowStr = tmp.replaceAll("(^[0-9]+):([0-9]+)$", "$1#$2");
		String values[];
		Integer pos1 = 0;
		Integer pos2 = 0;
		//行指定
		if (!tmp.equals(rowStr)) {
			values = rowStr.split("#");
			if (values.length == 2 ) {
				pos1 = Integer.valueOf(values[0]);
				pos2 = Integer.valueOf(values[1]);
				return new PoiRect(pos1-1,0,255,pos2-1);
			}
		}
		return null;
	}
	/**
	 * 
	 * @param value
	 * @return
	 */
	public static PoiRect analyseSelectColumnOrRow(String value) {
		if (value == null) return null;
		
		PoiRect result = null;
		
		result = analyseSelectColumn(value);
		if (result == null) {
			result = analyseSelectRow(value);
		}

		return result;
	}
	
	public static boolean equals(Object ...args) {
		Object lastItem = null;
		for (Object item : args) {
			if (lastItem != null) {
				if (!lastItem.equals(item) ) {
					return false;
				}
			}
			//前回値を保持
			lastItem = item;
		}
		return true;
	}
	
	public static PoiPosition analyseAddress(String value) {
		
		Integer x = 0;
		Integer y = 0;

		String col;
		String row;
		String tmp;
		
		if (value != null && value.length() >= 2) {
			//$は解除
			tmp = value.replaceAll("\\$", "");
			
			col = tmp.replaceAll("(^[A-Za-z]+)([0-9]+)$", "$1").toUpperCase();
			row = tmp.replaceAll("(^[A-Za-z]+)([0-9]+)$", "$2");
			
			if (col.equals(tmp.toUpperCase()) || row.equals(tmp)) {
				throw new ArrayIndexOutOfBoundsException("error[" + value + "]");
			}
			
			if (col.length() > 2 || col.length() == 0 || row.length() == 0) {
				throw new ArrayIndexOutOfBoundsException("error[" + value + "]");
			}
			
			if (col.length() >= 1) {
				x = alpabetToNumber(col);
			}
			
			y = Integer.valueOf(row) - 1;
			
			if (x >= 0 && x <= 255 && y >= 0 && y <= 65535) {
				return new PoiPosition(x, y);
			}
		}
		throw new ArrayIndexOutOfBoundsException("error[" + value + "]");
	}
	
	public static Integer alpabetToNumber(char x) {
		return x - 'A';
	}

	public static Integer alpabetToNumber(String value) {
		Integer result = 0;
		if (value != null ) {
			for(int i=0; i < value.length(); i++) {
				result *= 26;
				result += alpabetToNumber(value.charAt(i)) + 1;
			}
		}
		return result - 1;
	}
	/**
	 * 指定したsheetの最大座標を取得する
	 * @param sheet ワークシート
	 * @return 最大座標のポイント
	 */
	public static PoiPosition getEndPoint(HSSFSheet sheet) {
		//行

		Integer y = sheet.getLastRowNum();
		Integer x = 1;
		//列
		Iterator<Row> ite = sheet.rowIterator();
		Row item;
		int tmp = 1;
		while(ite.hasNext()) {
			item = ite.next();
			tmp = item.getLastCellNum();
			if (x < tmp ) {
				x = tmp;
			}
		}
		
		return new PoiPosition(x,y);
	}
	

	/**
	 * 指定したファイルを読み込み、バイト配列で返す
	 * @param f
	 * @return
	 */
	public static byte[] getByte(File f) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(f);
			return IOUtils.toByteArray(fis);
		}
		catch(Exception e){}
		finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (IOException e) {}
			}
		}
		return null;
	}
	/**
	 * シート存在確認
	 * @param name
	 * @return
	 */
	public static boolean exists(HSSFWorkbook book, String name) {
		return (book.getSheetIndex(name) != -1);
	}
	
	
	/**
	 * NULL VALUE
	 * @param <T>
	 * @param value
	 * @param value2
	 * @return
	 */
	public static <T> T nvl(T value, T value2) {
		if (value == null) {
			return value2;
		}
		return value;
	}
	
	public static void setSelectAndActive(HSSFSheet sheet) {
		HSSFWorkbook book = sheet.getWorkbook();
		//選択解除
		for (int i = 0 ; i < book.getNumberOfSheets(); i++ ) {
			book.getSheetAt(i).setSelected(false);
		}
		//選択
		sheet.setSelected(true);
		//アクティブ
		book.setActiveSheet( book.getSheetIndex(sheet) );
	}
	
	public static int getPictureType(File file) {
		if (file == null || file.getName() == null) return -1;
		return getPictureType(getSuffix(file.getName()).toLowerCase());
	}
	
	public static String getSuffix(String fileName ){
	    if (fileName == null)
	        return null;
	    int point = fileName.lastIndexOf(".");
	    if (point != -1) {
	        return fileName.substring(point + 1);
	    }
	    return fileName;
	}
	
	public static int getPictureType(String fileType) {
		if (fileType == null) return -1;
		
		String type = fileType.toLowerCase();
		if (TYPE_JPEG.equals(type) || TYPE_JPG.equals(type) ) {
			return HSSFWorkbook.PICTURE_TYPE_JPEG;
		}
		else if (TYPE_PNG.equals(type)) {
			return HSSFWorkbook.PICTURE_TYPE_PNG;
		}
		else if (TYPE_DIB.equals(type)) {
			return HSSFWorkbook.PICTURE_TYPE_DIB;
		}
		return -1;
	}
}
