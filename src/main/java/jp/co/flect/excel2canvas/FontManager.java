package jp.co.flect.excel2canvas;

import java.util.Map;
import java.util.HashMap;

public class FontManager {
	
	private static final Map<String, String> DEFAULT_MAP = new HashMap<String, String>();
	
	static {
		DEFAULT_MAP.put("メイリオ", "Meiryo");
		DEFAULT_MAP.put("ＭＳ Ｐゴシック", "Hiragino Kaku Gothic Pro");
		DEFAULT_MAP.put("ＭＳ Ｐ明朝", "Hiragino Mincho Pro");
		DEFAULT_MAP.put("ＭＳ ゴシック", "MS Gothic");
		DEFAULT_MAP.put("ＭＳ 明朝", "MS Mincho,");
	}
	
	private Map<String, String> map;
	
	public FontManager() {
		this.map =  new HashMap<String, String>(DEFAULT_MAP);
	}
	
	public FontManager(Map<String, String> map) {
		this();
		this.map.putAll(map);
	}
	
	public String getFontFamily(String fontName) {
		StringBuilder buf = new StringBuilder(fontName);
		String converted = this.map.get(fontName);
		if (converted != null) {
			buf.append(",").append(converted);
		}
		String ret = buf.toString();
		if (ret.indexOf("ゴシック") != -1 || ret.indexOf("Gothic") != -1 || ret.indexOf("ｺﾞｼｯｸ") != -1) {
			buf.append(",sans-serif");
		} else if (ret.indexOf("明朝") != -1 || ret.indexOf("Mincho") != -1) {
			buf.append(",serif");
		} else if (ret.indexOf("行書") != -1 || ret.indexOf("草書") != -1) {
			buf.append(",cursive");
		}
		ret = buf.toString();
		return ret;
	}
	
	public void addFontFamilyMapping(String fontName, String fontFamily) {
		this.map.put(fontName, fontFamily);
	}
}
