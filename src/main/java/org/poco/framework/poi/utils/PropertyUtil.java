package org.poco.framework.poi.utils;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class PropertyUtil {

	private static final String STR_GETTER = "get";
	private static final String STR_SETTER = "set";
	private static final int TYPE_GETTER = 0;
	private static final int TYPE_SETTER = 1;
	/**
	 * インスタンスの存在を不許可
	 */
	private PropertyUtil() {
	}

	/**
	 * プロパティ名からgetter/setterメソッド名を作成する 
	 * @param propertyName プロパティ名
	 * @param type 0 = getter / not 0 = setter
	 * @return メソッド名
	 */
	private static String getMethodName(String propertyName,int type) {
		if (propertyName == null) return null;
		String typeName;
		String methodName;
		if (type == 0) {
			typeName = STR_GETTER;
		}
		else {
			typeName = STR_SETTER;
		}
		//メソッド名（getter/setter+先頭大文字プロパティ名
		methodName = typeName + propertyName.substring(0,1).toUpperCase();
		if (propertyName.length() != 1) {
			methodName +=  propertyName.substring(1);
		}
		return methodName;
	}
	
	/**
	 * getter/setterメソッド名からプロパティ名を作成する
	 * @param methodName
	 * @return
	 */
	private static String getPropertyName(String methodName) {
		if (methodName == null) return null;
		String propertyName;
		if (methodName.startsWith(STR_GETTER) || methodName.startsWith(STR_SETTER) ) {
			propertyName = methodName.substring(3, 4).toLowerCase();
			if (propertyName.length() - STR_GETTER.length() != 1 ) {
				propertyName += methodName.substring(4);
			}
			return propertyName; 
		}
		else {
			return methodName;
		}
	}

    /**
     * 指定したクラスのプロパティを文字列で指定し、クラスの型を取得します<BR/>
     * 取得できない場合はnullを返します。
     * @param classInstance
     * @param propertyName
     * @return Class型
     * @throws NoSuchFieldException
     * @throws IllegalAccessException
     */
    public static Class<?> getPropertyType(Object classInstance, String propertyName) {
    	Class<?> result = null;
    	
    	//public Fieldの場合
    	try {
   			result = getType(classInstance, propertyName);
    	}
    	catch (Exception e) {
    		//エラーは無視
    	}
    	//フィールドで取得できなかった場合は、getter/setterを検索する
    	if (result == null) {
    		try {
    			Method method = classInstance.getClass().getMethod(getMethodName(propertyName,TYPE_GETTER));
   				result = method.getReturnType();
    		}
    		catch (Exception e){
    			//エラーは無視
    		}
    	}
    	return result;
    }

    /**
     * 指定したクラスのプロパティを文字列で指定し値を取得する
     *
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @return 値
     */
    public static Object getValue(Object classInstance, String propertyName) {
    	try {
        	//インスタンスからフィールドを取得し、値を取得する
    		return classInstance.getClass().getField(propertyName).get(classInstance);
    	}
    	catch (IllegalAccessException e) {
    	}
    	catch (NoSuchFieldException e) {
    	}
    	//getterメソッドから取得を試みる
    	try {
    		return classInstance.getClass().getMethod(getMethodName(propertyName,TYPE_GETTER)).invoke(classInstance);
    	}
    	catch (Exception e) {
    	}
    	return null;
    }

    /**
     * 強制的に指定したクラスのプロパティを文字列で指定し値を取得する<br/>
     * privateフィールドも取得可能
     *
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @return 値
     */
    public static Object getForceValue(Object classInstance, String propertyName) throws NoSuchFieldException,IllegalAccessException {
    	//インスタンスからフィールドを取得し、値を取得する
		Field nameField = classInstance.getClass().getDeclaredField(propertyName);
		nameField.setAccessible(true);
   		return nameField.get(classInstance);
    }

    /**
     * 指定したクラスのプロパティを文字列で指定し値を設定する
     *
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @param value 設定する値を指定
     */
    public static void setValue(Object classInstance, String propertyName, Object value) throws NoSuchFieldException,IllegalAccessException{
    	Class<?> type;
    	if (value == null) {
    		type = getPropertyType(classInstance,propertyName);
    	}
		else {
			type = value.getClass();
		}
		setValue(classInstance, propertyName, value, type);
    }
    
    /**
     * 指定したクラスのプロパティを文字列で指定し値を設定する
     *
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @param value 設定する値を指定
     * @param clazz setterのクラス型
     */
    public static void setValue(Object classInstance, String propertyName, Object value, Class<?> clazz) throws NoSuchFieldException,IllegalAccessException{
    	IllegalAccessException ie = null;
    	NoSuchFieldException ne = null;
    	try {
	    	//インスタンスからフィールドを取得し、値を設定する
	   		classInstance.getClass().getField(propertyName).set(classInstance,value);
	   		return;
    	}
    	catch (IllegalAccessException e) {
    		ie = e;
    	}
    	catch (NoSuchFieldException e) {
    		ne = e;
    	}
    	//setterメソッドへの設定を試みる
    	try {
    		if (clazz != null) {
    			classInstance.getClass().getMethod(getMethodName(propertyName,TYPE_SETTER),clazz).invoke(classInstance, value);
    		}
    		return;
    	}
    	catch (Exception e) {
    		if (ie != null) {
    			throw ie;
    		}
    		else {
	    		throw ne;
    		}
    	}
    }

	/**
	 * 指定したクラスのプロパティ一覧を取得する。（getter/setterは含まない）
	 *
	 * @param classInstance クラスのインスタンスを指定
	 * @return IPropertyインターフェースを実装したクラスを格納したList
	 */
	public static List<IProperty> getProperties(Object classInstance){
		Field[] list;
		IProperty prop = null;

		List<IProperty> _propertiesLisy_= new ArrayList<IProperty>();

		list = classInstance.getClass().getFields();

		for(int i=0;i<list.length;i++) {
			prop = createProperty();
			prop.setName(list[i].getName());
			prop.setType(list[i].getType().getName());
			_propertiesLisy_.add(prop);
		}

		return _propertiesLisy_;
	}

	/**
	 * 指定したクラスのプロパティ一覧を取得する<BR/>
	 * getter/setterもプロパティとして取得する
	 *
	 * @param Object クラスのインスタンスを指定
	 * @return IPropertyインターフェースを実装したクラスを格納したList
	 */
	public static List<IProperty> getPropertiesAll(Object classInstance){
		IProperty prop = null;
		List<IProperty> _propertiesList_= getProperties(classInstance);

		Method[] mlist = classInstance.getClass().getMethods();
		String propertyName;
		HashMap<String, Class<?>> methodMap = new HashMap<String, Class<?>>();
		//methodMapを作成する(getterはoverloadができない為、キーとする）
		for (Method mitem : mlist) {
			//getterを対象
			if (mitem.getName().startsWith(STR_GETTER)) {
				//引数がなく、戻り値がvoid以外のものをgetterと認識
				if (mitem.getParameterTypes().length == 0 && !mitem.getReturnType().getName().equals("void") ) {
					propertyName = getPropertyName(mitem.getName());
					methodMap.put(propertyName, mitem.getReturnType());
				}
			}
		}
		//プロパティを追加する
		for (Method mitem : mlist) {
			//setterを対象
			if (mitem.getName().startsWith(STR_SETTER) ) {
				//引数が1個の場合のみsetterと認識
				if (mitem.getParameterTypes().length != 1) {
					continue;
				}
				propertyName = getPropertyName(mitem.getName());
				prop = getPropertyItem(_propertiesList_, propertyName);
				//既にプロパティがあるものは除外
				if (prop != null) {
					continue;
				}
				//getterが存在しないものは除外
				if(!methodMap.containsKey(propertyName)) {
					continue;
				}
				//変数の型が一致しない場合は除外
				if (!methodMap.get(propertyName).equals(mitem.getParameterTypes()[0]) ) {
					continue;
				}
				//一致するものはプロパティと認識
				prop = createProperty();
				//プロパティ名設定
				prop.setName(propertyName);
				//型を設定
				prop.setType( methodMap.get(propertyName).getName() );
				//リストに追加
				_propertiesList_.add(prop);
 			}
		}
		
		return _propertiesList_;
	}

	/**
	 * 指定したリストから指定したプロパティ名が存在する場合は
	 * インスタンスを返す。存在しない場合は、nullを返す。
	 *
	 * @param list プロパティリスト
	 * @param propertyName プロパティ名
	 * @return IPropertyインスタンス又はnull
	 */
	private static IProperty getPropertyItem(List<IProperty> list, String propertyName) {
		for (IProperty item : list) {
			if (item.getName().equals(propertyName) ) {
				return item;
			}
		}
		return null;
	}

    /**
     * 指定したクラスのプロパティを文字列で指定しクラスを取得する
     *
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @return Class型
     */
    public static Class<?> getType(Object classInstance, String propertyName) throws NoSuchFieldException,IllegalAccessException {
    	//インスタンスからフィールドを取得し、値を取得する
   		return classInstance.getClass().getField(propertyName).getType();
    }

    /**
     * 強制的に指定したクラスのプロパティを文字列で指定しクラスを取得する<br/>
     * privateフィールドも取得可能
     *
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @return Class型
     */
    public static Class<?> getForceType(Object classInstance, String propertyName) throws NoSuchFieldException,IllegalAccessException {
    	//インスタンスからフィールドを取得し、値を取得する
		Field nameField = classInstance.getClass().getDeclaredField(propertyName);
		nameField.setAccessible(true);
   		return nameField.getType();
    }
    
    /**
     * 指定したクラスのプロパティが存在するかを確認する<BR/>
     * 速度を優先する場合は、getPropertiesAllで一覧を取得し、IPropertyのnameで存在確認をするべき。
     * @param classInstance クラスのインスタンスを指定
     * @param propertyName プロパティ名を指定
     * @return 
     */
    public static Boolean hasOwnProperty(Object classInstance, String propertyName) {
    	if (classInstance == null) {
    		return false;
    	}
    	try {
    		if (classInstance.getClass().getField(propertyName) != null) {
    			return true;
    		}
    	}
    	catch (NoSuchFieldException e) {
    		//getter/setterかもしれない
    	}
    	try {
    		//NULLの場合は、NoSuchMethodExceptionが上がるから、通ればOK
    		Method method = classInstance.getClass().getMethod(getMethodName(propertyName,TYPE_GETTER));
    		method = classInstance.getClass().getMethod(getMethodName(propertyName,TYPE_SETTER));
    		if (method != null) {
    			return true;
    		}
    	}
    	catch (NoSuchMethodException e) {
    		//ないみたい
    	}
    	return false;
    }
    
    public static IProperty createProperty() {
    	return new IProperty() {
    		private String _name;
    		private String _type;
    		public String getName() {
    			return _name;
    		}
    		public String getType() {
    			return _type;
    		}
    		public void setName(String value) {
    			_name = value;
    		}
    		public void setType(String value) {
    			_type = value;
    		}
    	};
    }
    
    public static interface IProperty {
    	/**
    	 * 結果を取得する
    	 * @return タイプ（型）
    	 */
    	public String getType();

    	/**
    	 * 結果を設定する
    	 * @param value 結果
    	 */
    	public void setType(String value);

    	/**
    	 * 結果を取得する
    	 * @return プロパティ名
    	 */
    	public String getName();

    	/**
    	 * 結果を設定する
    	 * @param value 結果
    	 */
    	public void setName(String value);
    }
}
