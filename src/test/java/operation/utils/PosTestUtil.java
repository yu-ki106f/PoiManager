package operation.utils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.poco.framework.poi.utils.PropertyUtil;
import org.poco.framework.poi.utils.PropertyUtil.IProperty;

import operation.dto.DetailDto;
import operation.dto.SampleParentDto;
import operation.dto.HeaderDataDto;

public class PosTestUtil {
	//---- sample data create ----
	public static SampleParentDto createData(Integer count) {
		SampleParentDto result = new SampleParentDto();
		result.header = new HeaderDataDto();
		result.header.date = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format(new Date());
		result.list = createList(count);
		return result;
	}
	private static List<DetailDto> createList(Integer count) {
		List<DetailDto> result = new ArrayList<DetailDto>();
		for (Integer i = 0; i < count; i++) {
			result.add(createDetail(i));
		}
		return result;
	}
	private static DetailDto createDetail(Integer value) {
		DetailDto result = new DetailDto();
		result.customerName = "customerName " + value.toString();
		result.customerSimpleName = "簡易顧客名" + value.toString();
		result.progressName = value.doubleValue() / 100;
		try {
			SimpleDateFormat df = new SimpleDateFormat("yyyyMMddHHmmss");
			Integer a = (value & 0xF) + 10;
			result.createDate = df.parse("201412"+a.toString()+"153015");
		} catch (ParseException e) {
		}
		result.reliabilityName = "確度" + value.toString();
		result.projectName = "プロジェクト名" + value.toString();
		result.remarks = "備考" + value.toString();
		return result;
	}
	/**
	 * リスト比較
	 * @param list1
	 * @param list2
	 * @return
	 */
	public static boolean equals(List<?> list1, List<?> list2){
		try {
			if (list1 == null && list2 == null) {
				return true;
			}
			if (list1.size() != list2.size()) return false;
			if (list1.size() == 0) return true;
			
			if (!equalsPropertys(list1.get(0),list2.get(0)) ) {
				return false;
			}
			for (Integer i = 0; i < list1.size(); i++) {
				if (!equals(list1.get(i),list2.get(i))) {
					return false;
				}
			}
			return true;
		}
		catch(NullPointerException e) {
			return false;
		}
	}
	/**
	 * プロパティ一致
	 * @param dto1
	 * @param dto2
	 * @return
	 */
	public static boolean equalsPropertys(Object dto1, Object dto2) {
		List<IProperty> prop1 = PropertyUtil.getPropertiesAll(dto1);
		List<IProperty> prop2 = PropertyUtil.getPropertiesAll(dto2);
		if (prop1.size() != prop2.size()) return false;
		IProperty obj1;
		IProperty obj2;
		for (Integer i=0; i<prop1.size();i++) {
			obj1 = prop1.get(i);
			obj2 = prop2.get(i);
			if ( !obj1.getName().equals(obj2.getName())) {
				return false;
			}
			if ( !(obj1.getType().equals(obj2.getType()) ) ) {
				return false;
			}
		}
		return true;
	}
	/**
	 * オブジェクト一致
	 * @param dto1
	 * @param dto2
	 * @return
	 */
	public static boolean equals(Object dto1, Object dto2) {
		List<IProperty> prop1 = PropertyUtil.getPropertiesAll(dto1);
		Object obj1;
		Object obj2;
		for (Integer i = 0; i < prop1.size(); i++) {
			obj1 = PropertyUtil.getValue(dto1, prop1.get(i).getName());
			obj2 = PropertyUtil.getValue(dto2, prop1.get(i).getName());
			if (isEmptyEquals(obj1,obj2)) continue;
			if (obj1 == null || obj2 == null) {
				return false;
			}
			if (!obj1.equals(obj2)) {
				return false;
			}
		}
		return true;
	}
	/**
	 * NULLと空文字判定
	 * @param a
	 * @param b
	 * @return
	 */
	private static boolean isEmptyEquals(Object a, Object b) {
		if (a != null && b != null) return false;
		if (a == null && b == null) return true;
		Object notNull = (a == null) ? b : a;
		if ("".equals(notNull.toString())) {
			return true;
		}
		return false;
	}
}
