package org.poco.framework.poi.managers;

import java.io.File;
import java.util.List;

import org.poco.framework.poi.converter.IValueConverter;
import org.poco.framework.poi.creator.comment.ICommentCreator;
import org.poco.framework.poi.creator.image.IImageCreator;
import org.poco.framework.poi.dto.PoiPosition;
import org.poco.framework.poi.dto.PoiRect;
import org.poco.framework.poi.dto.PoiStyleDto;
import org.poco.framework.poi.exception.PoiException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public interface IPoiManager {

	public interface IPoiBook {

		/**
		 * 指定した名前のシートを新しい名前のシートにコピーします。
		 * @param fromName
		 * @param toName
		 * @return
		 */
		IPoiSheet cloneSheet(String fromName, String toName);

		/**
		 * ない場合は作成します。
		 * @param name
		 * @return
		 */
		IPoiSheet sheet(String name);

		/**
		 * Bookを保存します。
		 * @param e
		 * @return
		 */
		boolean saveBook(File f);

		/**
		 * 内部に指定されているファイルにBookを保存します。
		 * @param e
		 * @return
		 */
		boolean saveBook();
		/**
		 * readwriter取得
		 * @return
		 */
		IReadWriter getReadWriter();

		/**
		 * オリジナルWorkBook取得
		 * @return
		 */
		Workbook getOrgWorkBook();
	}

	public interface IPoiSheet extends IPoiBook {
		/**
		 * オリジナルのシートを返す
		 * @param name
		 * @return
		 */
		Sheet getOrgSheet();
		IPoiBook getBook();

		IPoiCell cell(int x, int y);
		IPoiCell cell(PoiPosition pos);
		IPoiCell cell(String address);

		IPoiRange range(String address);
		IPoiRange range(PoiRect rect);
		IPoiRange range(int fromX, int fromY, int toX, int toY);

		IPoiSheet setSheetName(String name);

		IPoiRange merge(String address);
		IPoiRange merge(PoiRect rect);
		IPoiRange merge(int fromX, int fromY, int toX, int toY);

		IPoiRange insertRows(int insertPos);
		IPoiRange insertRows(int insertPos, int insertSize);
		IPoiRange insertRows(String address);

		IPoiSheet selected();

		String sheetName();

	}

	public interface IPoiRange extends IPoiSheet {
		IPoiCell[] getCells();

		Boolean isMerged();

		List<Object> getValue();
		List<String> getStringValue();

		IPoiRange setValue(Object value);
		<T> IPoiRange setValue(Object value, Class<T> clazz);

		IImageCreator<IPoiRange> setImage(File f);
		IImageCreator<IPoiRange> setImage(byte[] bytes, String fileName);
		ICommentCreator<IPoiRange> comment(String comment);

		IPoiRange setColumnWidth(double point);
		IPoiRange setRowWidth(double point);

		IPoiRange setColumnHidden(boolean hidden);
		IPoiRange setRownHidden(boolean hidden);

		IPoiRange autoSizeColumn();
		IPoiRange autoSizeColumn(boolean useMergedCells);

		IStyleManager<IPoiRange> style();
		IStyleManager<IPoiRange> style(PoiStyleDto dto);
	}

	public interface IPoiCell extends IPoiSheet {

		Row getOrgRow();
		Cell getOrgCell();

		IValueConverter value();

		String getStringValue();
		Object getValue();


		IPoiCell setValue(Object value);
		<T> IPoiCell setValue(Object value, Class<T> clazz);
		IPoiCell setFormura(String value);
		IPoiCell setError(Byte value);

		IImageCreator<IPoiCell> setImage(File f);
		IImageCreator<IPoiCell> setImage(byte[] bytes, String fileName);
		ICommentCreator<IPoiCell> comment(String comment);

		IPoiCell setColumnWidth(double point);
		IPoiCell setRowWidth(double point);

		IPoiCell setColumnHidden(boolean hidden);
		IPoiCell setRownHidden(boolean hidden);

		IPoiCell autoSizeColumn();
		IPoiCell autoSizeColumn(boolean useMergedCells);

		Integer x();
		Integer y();

		IStyleManager<IPoiCell> style();
		IStyleManager<IPoiCell> style(PoiStyleDto dto);

		IObjectWriter writeObject(List<Object> list) throws PoiException;
		IObjectWriter writeObject(Object dto) throws PoiException;
	}

	public interface IObjectWriter {
		IPoiCell write() throws PoiException;
		IObjectWriter setHeader(Object header);
		IObjectWriter order(List<String> propertiesList) throws PoiException;
		IObjectWriter order(String[] propertiesArray) throws PoiException;
	}

}
