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
import com.google.gson.JsonDeserializer;
import com.google.gson.JsonDeserializationContext;
import com.google.gson.JsonParseException;
import com.google.gson.JsonElement;
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
	private transient Map<String, StrInfo> strMap = null;
	
	public int getWidth() { return this.width;}
	public void setWidth(int n) { this.width = n;}
	
	public int getHeight() { return this.height;}
	public void setHeight(int n) { this.height = n;}
	
	public void addLineInfo(LineInfo l) { this.lines.add(l);}
	public void addFillInfo(FillInfo f) { this.fills.add(f);}
	public void addStrInfo(StrInfo s) { this.strs.add(s);}
	public void addPictureInfo(PictureInfo p) { this.pictures.add(p);}
	public void addChartInfo(ChartInfo p) { this.charts.add(p);}
	
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
	
	public String toJson() {
		return new Gson().toJson(this);
	}
	
	public static ExcelToCanvas fromJson(String json) {
		GsonBuilder builder = new GsonBuilder();
		builder.registerTypeAdapter(Chart.class, new ChartDeserializer());
		return builder.create().fromJson(json, ExcelToCanvas.class);
	}
	
	public static class FillInfo {
		
		private int[] p;
		private String back;
		private String fore;
		private Integer pattern;
		
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
		
		public int[] getPoints() { return this.p;}
		public String getBackground() { return this.back;}
		public String getForeground() { return this.fore;}
		public int getPattern() { return this.pattern == null ? 0 : this.pattern.intValue();}
		
	}
	
	private static String mapToStyle(Map<String, String> map) {
		StringBuilder buf = new StringBuilder();
		for (Map.Entry<String, String> entry : map.entrySet()) {
			buf.append(entry.getKey()).append(":")
				.append(entry.getValue()).append(";");
		}
		return buf.toString();
	}
	
	private static Map<String, String> styleToMap(String str) {
		Map<String, String> map = new HashMap<String, String>();
		String[] values = str.split(";");
		for (String value : values) {
			value = value.trim();
			int idx = value.indexOf(':');
			map.put(value.substring(0, idx).trim(), value.substring(idx + 1).trim());
		}
		return map;
	}
	
	public static class StrInfo {
		
		private int[] p;
		private String id;
		private String text;
		private String align;
		private String style;
		private String link;
		private Boolean formula;
		private String rawdata;
		private String comment;
		private Integer commentWidth;
		private transient Map<String, String> styleMap;
		private String clazz;
		
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
		public String getStyle() { return this.style;}
		public String getLink() { return this.link;}
		public String getComment() { return this.comment;}
		public int getCommentWidth() { return this.commentWidth != null ? this.commentWidth.intValue() : 0;}
		public boolean isFormula() { return this.formula != null && this.formula.booleanValue();}
		public void setFormula(boolean b) { this.formula = b ? Boolean.TRUE : null;}
		
		public Map<String, String> getStyleMap() { return this.styleMap;}
		
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
		
		public void resetStyle(String name, String value) {
			if (this.styleMap == null) {
				this.styleMap = styleToMap(this.style);
			}
			this.styleMap.put(name, value);
			this.style = mapToStyle(this.styleMap);
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
	
	private static class ChartDeserializer implements JsonDeserializer<Chart> {
		public Chart deserialize(JsonElement json, Type typeOfT, JsonDeserializationContext context) throws JsonParseException {
			return context.deserialize(json, Flotr2.class);
		}
	}
	
}
