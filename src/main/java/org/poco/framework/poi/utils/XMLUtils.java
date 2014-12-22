package org.poco.framework.poi.utils;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class XMLUtils {

	/**
	 * Nodeからすべての属性を取得する
	 * @param node XMLノード
	 * @return 属性リスト
	 */
	public static List<Map<String,String>> getAttributeList(Node node) {

		// 引数チェック
		if(node == null) {
			return null;
		}
		
		// 戻値
		List<Map<String,String>> result = new ArrayList<Map<String,String>>();
		Map<String,String> item = null;
		
		// ノードの属性を取得
		NamedNodeMap nodeMap = node.getAttributes();
		
		if(nodeMap != null) {
			
			for(int i = 0; i < nodeMap.getLength(); i++) {
				
				Node work =  nodeMap.item(i);
				item = new HashMap<String, String>();
				
				// 属性の名前
				item.put("name",work.getNodeName());
				
				// 属性の値
				item.put("value",work.getNodeValue());
				
				// 結果に追加
				result.add(item);
				
			}
		}
		
		return result;
	}
	
	/**
	 * Nodeから属性名を指定して値を取得する
	 * @param node XMLノード
	 * @param name 属性名
	 * @return 属性値
	 */
	public static String getAttributeValue(Node node, String name) {

		// 引数チェック
		if(node == null || name == null) {
			return null;
		}
		
		// 戻値
		String result = null;
		
		// ノードの属性を取得
		Node attribute = node.getAttributes().getNamedItem(name);
		
		if(attribute != null) {
		
			// 属性の値を取得
			result = attribute.getNodeValue();
		}
		
		return result;
		
	}
	
	/**
	 * Nodeから要素(#text)を取得する
	 * @param node XMLノード
	 * @return textの内容
	 */
	public static String getText(Node node) {

		// 引数チェック
		if(node == null) {
			return null;
		}
		
		// 戻値
		String result = null;
		
		// ノードから子要素を取得する
		NodeList nodeList = node.getChildNodes();
		
		for(int i = 0; i < nodeList.getLength(); i++) {
		
			Node value = nodeList.item(i);
		
			if(value.getNodeType() == Node.TEXT_NODE && 
				value.getNodeValue().trim().length() > 0){
				
				// テキスト要素発見
				result = value.getNodeValue();
				continue;
				
			}
		}

		return result;
	}
	
	/**
	 * Nodeからすべての直属の子Nodeのみ取得する
	 * @param node XMLノード
	 * @return 属性リスト
	 */
	public static List<Node> getChildNodeList(Node node) {

		// 引数チェック
		if(node == null) {
			return null;
		}
		
		// 戻値
		List<Node> result = new ArrayList<Node>();
		
		// list
		NodeList nodeList = node.getChildNodes();
		
		for(int i = 0; i < nodeList.getLength(); i++) {
		
			Node item = nodeList.item(i);
			
			// 要素の場合
			if(item.getNodeType() == Node.ELEMENT_NODE) {
				
				// 結果に追加
				result.add(item);
			}
		}
		
		return result;
	}
	
	/**
	 * Nodeから要素名を指定してNodeを取得する
	 * @param node XMLノード
	 * @param name 属性名
	 * @return Node
	 */
	public static Node getChildNode(Node node, String name) {

		// 引数チェック
		if(node == null || name == null) {
			return null;
		}
		
		// list
		NodeList nodeList = node.getChildNodes();
		
		for(int i = 0; i < nodeList.getLength(); i++) {
		
			Node item = nodeList.item(i);
			
			// 要素の場合
			if(item.getNodeName().equals(name)) {
				
				// 結果に追加
				return item;
			}
		}
		
		return null;
	}
	
	/**
	 * Nodeから要素名を指定して要素(#text)を取得する
	 * @param node XMLノード
	 * @param name 属性名
	 * @return textの内容
	 */
	public static String getChildNodeText(Node node, String name) {

		// 引数チェック
		if(node == null || name == null) {
			return null;
		}
		
		// 戻値
		String result = null;
		
		// list
		NodeList nodeList = node.getChildNodes();
		
		for(int i = 0; i < nodeList.getLength(); i++) {
		
			Node item = nodeList.item(i);
			
			// 要素の場合
			if(item.getNodeName().equals(name)) {
				
				// 結果に追加
				result = getText(item);
			}
		}
		
		return result;
	}
}
