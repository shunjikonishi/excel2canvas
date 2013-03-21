package jp.co.flect.excel2canvas;

import java.awt.Color;
import java.text.DecimalFormat;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Date;
import java.math.BigDecimal;

/**
 * Get cell value as HTML string.
 */
public class FormattedValue {
	
	public enum Type {
		BLANK,
		NUMBER,
		DATE,
		BOOLEAN,
		STRING,
		FORMULA,
		ERROR
	};
	
	public static final FormattedValue EMPTY = new FormattedValue("", Type.BLANK, "");
	
	private String value;
	private Color color;
	private Object rawdata;
	private Type type;
	private String formula;
	
	public FormattedValue(String value, Type type, Object rawdata) {
		this(value, type, rawdata, null);
	}
	
	public FormattedValue(String value, Type type, Object rawdata, Color color) {
		this.value = value;
		this.type = type;
		this.rawdata = rawdata;
		this.color = color;
	}
	
	public String getValue() { return this.value;}
	public void setValue(String s) { this.value = s;}
	
	public Type getType() { return this.type;}
	public void setType(Type t) { this.type = t;}
	
	public Object getRawData() { return this.rawdata;}
	public void setRawData(Object o) { this.rawdata = o;}
	
	public String getFormula() { return this.formula;}
	public void setFormula(String s) { this.formula = s;}
	
	public Color getColor() { return this.color;}
	public void setColor(Color c) { this.color = c;}
	
	public String toString() {
		if (this.color == null) return this.value;
		
		return new StringBuilder()
			.append("[").append(getColorString(this.color))
			.append("]").append(this.value)
			.toString();
	}
	
	private static StringBuilder appendNumber(StringBuilder buf, int n) {
		if (n < 10) {
			buf.append("0");
		}
		buf.append(n);
		return buf;
	}
	
	public String getRawString() {
		if (this.rawdata == null) {
			return null;
		} else if (this.rawdata instanceof Date) {
			Calendar cal = new GregorianCalendar();
			cal.setTime((Date)this.rawdata);
			int year = cal.get(Calendar.YEAR);
			int month = cal.get(Calendar.MONTH) + 1;
			int day = cal.get(Calendar.DATE);
			int hour = cal.get(Calendar.HOUR_OF_DAY);
			int min = cal.get(Calendar.MINUTE);
			int sec = cal.get(Calendar.SECOND);
			
			StringBuilder buf = new StringBuilder();
			if (year != 1900 || month == 1 || day == 1) {
				buf.append(year).append("-");
				appendNumber(buf, month).append("-");
				appendNumber(buf, day);
			}
			if (hour != 0 || min != 0 || sec != 0) {
				if (buf.length() > 0) {
					buf.append(" ");
				}
				appendNumber(buf, hour).append(":");
				appendNumber(buf, min).append(":");
				appendNumber(buf, sec);
			}
			return buf.toString();
		} else if (this.rawdata instanceof Double) {
			double d = ((Double)this.rawdata).doubleValue();
			if (Math.floor(d) == d) {
				BigDecimal bd = new BigDecimal(d);
				bd.setScale(0);
				return bd.toString();
			} else {
				return Double.toString(d);
			}
		} else {
			return this.rawdata.toString();
		}
	}
	
	public String getClorString() {
		return getColorString(this.color);
	}
	
	public static String getColorString(Color color) {
		if (color == null) {
			return null;
		}
		int a = color.getAlpha();
		int r = color.getRed();
		int g = color.getGreen();
		int b = color.getBlue();
		StringBuilder buf = new StringBuilder();
		if (a == 255) {
			buf.append("#")
				.append(toHex(r))
				.append(toHex(g))
				.append(toHex(b));
		} else {
			buf.append("rgba(")
				.append(r).append(",")
				.append(g).append(",")
				.append(b);
			double d = 1.0 / 255 * a;
			buf.append(new DecimalFormat("0.##").format(d));
			buf.append(")");
		}
		return buf.toString();
	}
	
	private static String toHex(int n) {
		String s = Integer.toHexString(n).toUpperCase();
		return s.length() == 1 ? "0" + s : s;
	}
	
}