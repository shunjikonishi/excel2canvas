package jp.co.flect.excel2canvas;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.Iterator;
import java.lang.reflect.Type;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.PictureData;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.commons.codec.binary.Base64;
import jp.co.flect.excel2canvas.chart.Chart;
import jp.co.flect.excel2canvas.chart.Flotr2;

/**
 * This class is able to convert to JSON string.
 * It is able to draw on the web browser by jquery.excel2canvas.js (depend on HTML5 canvas)
 */
public class ExcelToCanvas {
	
	private int width;
	private int height;
	private List<LineInfo> lines = new ArrayList<LineInfo>();
	private List<FillInfo> fills = new ArrayList<FillInfo>();
	private List<StrInfo> strs = new ArrayList<StrInfo>();
	private List<PictureInfo> pictures = new ArrayList<PictureInfo>();
	private List<ChartInfo> charts = new ArrayList<ChartInfo>();
	
	private transient int styleIndex = 0;
	private Map<String, String> styles = new HashMap<String, String>();
	
	private transient Map<String, StrInfo> strMap = null;
	
	public int getWidth() { return this.width;}
	public void setWidth(int n) { this.width = n;}
	
	public int getHeight() { return this.height;}
	public void setHeight(int n) { this.height = n;}
	
	public void addLineInfo(LineInfo l) { this.lines.add(l);}
	public void addPictureInfo(PictureInfo p) { this.pictures.add(p);}
	public void addChartInfo(ChartInfo p) { this.charts.add(p);}
	
	public void addFillInfo(FillInfo f) { 
		this.fills.add(f);
		f.setParent(this);
	}
	
	public void addStrInfo(StrInfo s) { 
		this.strs.add(s);
		s.setParent(this);
	}
	
	public List<LineInfo> getLines() { return this.lines;}
	public void setLines(List<LineInfo> list) { this.lines = list;}
	
	public List<FillInfo> getFills() { return this.fills;}
	public void setFills(List<FillInfo> list) { this.fills = list;}
	
	public List<StrInfo> getStrs() { return this.strs;}
	public void setStrs(List<StrInfo> list) { this.strs = list;}
	
	public List<PictureInfo> getPictures() { return this.pictures;}
	public void setPictures(List<PictureInfo> list) { this.pictures = list;}
	
	public List<ChartInfo> getCharts() { return this.charts;}
	public void setCharts(List<ChartInfo> list) { this.charts = list;}
	
	public StrInfo getStrInfo(String id) {
		if (this.strMap == null) {
			this.strMap = new HashMap<String, StrInfo>();
			for (StrInfo str : this.strs) {
				this.strMap.put(str.getId(), str);
			}
		}
		return this.strMap.get(id);
	}
	
	public void removeEmptyStrings() {
		Iterator<StrInfo> it = this.strs.iterator();
		while (it.hasNext()) {
			StrInfo str = it.next();
			if (str.getText() == null || str.getText().length() == 0) {
				it.remove();
			}
		}
		this.strMap = null;
	}
	
	private String addStyle(String str) {
		for (Map.Entry<String, String> entry : this.styles.entrySet()) {
			if (entry.getValue().equals(str)) {
				return entry.getKey();
			}
		}
		String key = "s" + (++styleIndex);
		styles.put(key, str);
		return key;
	}
	
	private String getStyle(String key) { return this.styles.get(key);}

	public Map<String, String> getStyles() { return this.styles;}
	
	//package local constructor
	ExcelToCanvas() {}
	
	public static class LineInfo {
		
		private int[] p;
		private Integer kind;//サイズ削減のためBORDER_THINはnullとする
		private String color;
		
		public LineInfo(int sx, int sy, int ex, int ey, int kind, String color) {
			this.p = new int[4];
			this.p[0] = sx;
			this.p[1] = sy;
			this.p[2] = ex;
			this.p[3] = ey;
			if (kind != CellStyle.BORDER_THIN) {
				this.kind = kind;
			}
			if (!"#000000".equals(color)) {
				this.color = color;
			}
		}
		
		public int[] getPoints() { return this.p;}
		public int getKind() { return this.kind == null ? CellStyle.BORDER_THIN : this.kind.intValue();}
		public String getColor() { return this.color;}
	}
	
	void loadInit() {
		if (this.fills != null) {
			for (FillInfo f : this.fills) {
				f.parent = this;
			}
		}
		if (this.strs != null) {
			for (StrInfo s : this.strs) {
				s.parent = this;
			}
		}
	}
	
	public String toJson() {
		return toJson(false);
	}

	public String toJson(boolean indent) {
		GsonBuilder builder = new GsonBuilder();
		if (indent) {
			builder = builder.setPrettyPrinting();
		}
		return builder.create().toJson(this);
	}

	public static ExcelToCanvas fromJson(String json) {
		GsonBuilder builder = new GsonBuilder();
		builder.registerTypeAdapter(Chart.class, new Flotr2());
		ExcelToCanvas excel = builder.create().fromJson(json, ExcelToCanvas.class);
		excel.loadInit();
		return excel;
	}
	
	public static class FillInfo {
		
		private transient ExcelToCanvas parent = null;
		
		private int[] p;
		private /* transient */ String back;//for compatibility
		private /* transient */ String fore;//for compatibility
		private /* transient */ Integer pattern;//for compatibility
		private String styleRef;
		
		public FillInfo(int sx, int sy, int ex, int ey, String back, String fore, int pattern) {
			this.p = new int[4];
			this.p[0] = sx;
			this.p[1] = sy;
			this.p[2] = ex;
			this.p[3] = ey;
			this.back = back;
			this.fore = fore;
			if (pattern > 1) {
				this.pattern = pattern;
			}
		}
		
		private void setParent(ExcelToCanvas excel) {
			this.parent = excel;
			
			StringBuilder buf = new StringBuilder();
			if (this.back != null) {
				buf.append(this.back);
			}
			buf.append("|");
			if (this.fore != null) {
				buf.append(this.fore);
			}
			buf.append("|");
			if (this.pattern != null) {
				buf.append(this.pattern);
			}
			this.styleRef = excel.addStyle(buf.toString());
			this.back = null;
			this.fore = null;
			this.pattern = null;
		}
		
		public int[] getPoints() { return this.p;}
		public String getBackground() { return this.parent == null ? this.back : getRef(0);}
		public String getForeground() { return this.parent == null ? this.fore : getRef(1);}
		public int getPattern() { 
			if (this.parent == null) {
				return this.pattern == null ? 0 : this.pattern.intValue();
			} else {
				String ref = getRef(2);
				return ref == null ? 0 : Integer.parseInt(ref);
			}
		}
		
		private String getRef(int idx) {
			String style = this.parent.getStyle(this.styleRef);
			String ret = style.split("|")[idx];
			return ret == null || ret.length() == 0 ? null : ret;
		}
		
	}
	
	public static String mapToStyle(Map<String, String> map) {
		StringBuilder buf = new StringBuilder();
		for (Map.Entry<String, String> entry : map.entrySet()) {
			buf.append(entry.getKey()).append(":")
				.append(entry.getValue()).append(";");
		}
		return buf.toString();
	}
	
	public static Map<String, String> styleToMap(String str) {
		Map<String, String> map = new HashMap<String, String>();
		String[] values = str.split(";");
		for (String value : values) {
			value = value.trim();
			int idx = value.indexOf(':');
			if (idx != -1) {
				map.put(value.substring(0, idx).trim(), value.substring(idx + 1).trim());
			}
		}
		return map.size() == 0 ? null : map;
	}
	
	public static class StrInfo {
		
		private transient ExcelToCanvas parent = null;
		
		private int[] p;
		private String id;
		private String text;
		private String align;
		private String styleRef;
		private String link;
		private Boolean formula;
		private String rawdata;
		private String comment;
		private /* transient */ String style;//for compatibility
		private Integer commentWidth;
		private transient Map<String, String> styleMap;
		private String clazz;
		private Map<String, String> dataAttrs = null;
		
		public StrInfo(int[] p, String id, String text, String align, Map<String, String> styleMap, String link, String comment, Integer commentWidth, boolean formula) {
			this.p = p;
			this.id = id;
			this.text = text;
			this.align = align;
			this.styleMap = styleMap;
			this.style = mapToStyle(styleMap);
			this.link = link;
			this.comment = comment;
			this.commentWidth = commentWidth;
			if (formula) {
				this.formula = Boolean.TRUE;
			}
		}
		
		
		public int[] getPoints() { return this.p;}
		public String getId() { return this.id;}
		public String getText() { return this.text;}
		
		public String getAlign() { return this.align;}
		public String getLink() { return this.link;}
		public String getComment() { return this.comment;}
		public int getCommentWidth() { return this.commentWidth != null ? this.commentWidth.intValue() : 0;}
		public boolean isFormula() { return this.formula != null && this.formula.booleanValue();}
		public void setFormula(boolean b) { this.formula = b ? Boolean.TRUE : null;}
		
		
		public String getRawData() { return this.rawdata;}
		public void setRawData(String s) { this.rawdata = s;}
		
		public String getClazz() { return this.clazz;}
		public void setClazz(String s) { this.clazz = s;}
		
		public void setText(String s) { 
			this.text = s;
			if (isAlignGeneral()) {
				char alignChar = isNumberString(s) ? 'r' : 'l';
				this.align = alignChar + this.align.substring(1);
			}
		}
		
		public void setText(String s, boolean alignLeft) {
			this.text = s;
			char alignChar = alignLeft ? 'l' : 'r';
			this.align = alignChar + this.align.substring(1);
		}
		
		public boolean isAlignGeneral() {
			return this.align != null && this.align.length() == 3 && this.align.charAt(2) == 'g';
		}
		
		private boolean isNumberString(String s) {
			if (s == null || s.length() == 0) {
				return false;
			}
			boolean dot = false;
			for (int i=0; i<s.length(); i++) {
				char c = s.charAt(i);
				if (c >= '0' && c <= '9') {
					continue;
				} else if (c == '-' && i == 0) {
					continue;
				} else if (c == ',') {
					continue;
				} else if (c == '.') {
					if (dot) {
						return false;
					}
					dot = true;
				} else {
					return false;
				}
			}
			return true;
		}
		
		private void setParent(ExcelToCanvas excel) {
			this.parent = excel;
			this.styleRef = excel.addStyle(this.style);
			this.style = null;
		}
		
		public String getStyle() {
			if (this.parent != null && this.styleRef != null) {
				return this.parent.getStyle(this.styleRef);
			} else {
				return this.style;
			}
		} 
		
		private void setStyle(String s) {
			if (this.parent == null) {
				this.style = s;
			} else {
				this.styleRef = this.parent.addStyle(s);
			}
		}
		
		public Map<String, String> getStyleMap() { return this.styleMap;}
		public void resetStyle(String name, String value) {
			if (this.styleMap == null) {
				this.styleMap = styleToMap(getStyle());
			}
			this.styleMap.put(name, value);
			setStyle(mapToStyle(this.styleMap));
		}

		public Map<String, String> getDataAttrs() { return this.dataAttrs;}
		public String getDataAttr(String name) { return this.dataAttrs == null ? null : this.dataAttrs.get(name);}
		public void setDataAttr(String name, String value) {
			if (this.dataAttrs == null) {
				this.dataAttrs = new HashMap<String, String>();
			}
			this.dataAttrs.put(name, value);
		}
	}
	
	public static class PictureInfo {
		
		private int[] p;
		private String uri;
		private String border;
		
		private transient byte[] data;
		private transient String mimeType;
		private transient String ext;
		
		public PictureInfo(PictureData data, int[] p, String border) {
			StringBuilder buf = new StringBuilder();
			buf.append("data:")
				.append(data.getMimeType())
				.append(";base64,");
			buf.append(Base64.encodeBase64String(data.getData()));
			this.uri = buf.toString();
			this.p = p;
			this.border = border;
			
			this.data = data.getData();
			this.mimeType = data.getMimeType();
			this.ext = data.suggestFileExtension();
		}
		
		public int[] getPoints() { return this.p;}
		public String getUri() { return this.uri;}
		public String getBorder() { return this.border;}
		public String getMimeType() { return this.mimeType;}
		public byte[] getData() { return this.data;}
		public String getExt() { return this.ext;}
	}
	
	public static class ChartInfo {
		
		private int[] p;
		private Chart chart;
		
		public ChartInfo(int[] p, Chart chart) {
			this.p = p;
			this.chart = chart;
		}
		
		public int[] getPoints() { return this.p;}
		public Chart getChart() { return this.chart;}
	}
	
}
