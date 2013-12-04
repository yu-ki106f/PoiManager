package org.poco.framework.poi.factory;

import java.util.HashMap;
import java.util.Map;

import org.poco.framework.poi.dto.PoiPosition;

public class PosFactory {

	private static Map<String,PoiPosition> map  = new HashMap<String, PoiPosition>();

	/**
	 * ポジションインスタンスを取得します
	 * @param x
	 * @param y
	 * @return
	 */
	public static synchronized PoiPosition getInstance(int x, int y) {
		//クラス参照では、ハッシュコードが変わってしまう為、文字とする
		String key = createKey(x,y);
		//取り出し
		PoiPosition result = map.get(key);
		if (result == null) {
			result = new PoiPosition(x,y);
			map.put(key, result);
		}
		return result;
	}

	/**
	 * 初期化
	 */
	public static void clear() {
		map.clear();
	}
	
	private static String createKey(int x, int y) {
		return String.valueOf(x).concat(":").concat(String.valueOf(y));
	}

}
