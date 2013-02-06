package jp.co.flect.excel2canvas.format;

import java.awt.Color;
import java.util.Locale;
import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.text.Format;

import org.apache.poi.ss.usermodel.Workbook;
import jp.co.flect.excel2canvas.ExcelColor;

/**
 * Wrapper class of Excel format information.
 */
public class FormatInfo {
	
	private static final Map<String, Locale> localeMap = new HashMap<String, Locale>();
	private static final Map<String, Color> colorMap = new HashMap<String, Color>();
	
	static {
		localeMap.put("$-411", new Locale("ja", "JP", "JP"));
		
		colorMap.put("black", Color.black);
		colorMap.put("blue", Color.blue);
		colorMap.put("cyan", Color.cyan);
		colorMap.put("green", Color.green);
		colorMap.put("magenta", Color.magenta);
		colorMap.put("red", Color.red);
		colorMap.put("white", Color.white);
		colorMap.put("yellow", Color.yellow);
	}
	
	public static FormatInfo[] parse(Workbook workbook, String str) {
		StringBuilder buf = new StringBuilder();
		Locale locale = null;
		Color color = null;
		Condition condition = null;
		
		List<FormatInfo> list = new ArrayList<FormatInfo>();
		
		int len = str.length();
		for (int i=0; i<len; i++) {
			char c = str.charAt(i);
			switch (c) {
				case ';':
					list.add(new FormatInfo(buf.toString(), locale, color, condition));
					buf.setLength(0);
					locale = null;
					color = null;
					condition = null;
					break;
				case '\\':
				case '_':
				case '*':
					buf.append(c);
					if (i+1 <len) {
						buf.append(str.charAt(i+1));
						i++;
					}
					break;
				case '"':
				{
					int idx = str.indexOf('"', i+1);
					if (idx != -1) {
						buf.append(str.substring(i, idx+1));
						i = idx;
					} else {
						buf.append(c);
					}
					break;
				}
				case '[':
				{
					int idx = str.indexOf(']', i+1);
					if (idx != -1) {
						boolean append = false;
						String key = str.substring(i+1, idx);
						if (key.length() > 0) {
							char c2 = key.charAt(0);
							switch (c2) {
								case '$'://Locale
								{
									locale = localeMap.get(key);
									if (locale == null) {
										locale = Locale.getDefault();
									}
									break;
								}
								case '>':
								case '<':
								case '=':
									condition = createCondition(key);
									break;
								default:
									if ((c2 == 'h' || c2 == 'm' || c2 == 's') &&
									    (key.length() == 1 || (key.length() == 2 && key.charAt(1) == c2)))
									{
										append = true;
									} else if (key.startsWith("Color")) {
										try {
											int colorIndex = Integer.parseInt(key.substring(5)) + 7;
											color = new ExcelColor(workbook).getColor((short)colorIndex);
										} catch (NumberFormatException e) {
											//Ignore
										}
									} else {
										color = colorMap.get(key.toLowerCase());
										if (color == null) {
											append = true;
										}
									}
									break;
							}
						}
						if (append) {
							buf.append(str.substring(i, idx+1));
						}
						i = idx;
					} else {
						buf.append(c);
					}
					break;
				}
				default:
					buf.append(c);
					break;
			}
		}
		if (buf.length() > 0) {
			list.add(new FormatInfo(buf.toString(), locale, color, condition));
		}
		if (list.size() > 1) {
			for (int i=0; i<list.size(); i++) {
				FormatInfo f = list.get(i);
				if (f.getCondition() == null) {
					switch (i) {
						case 0:
							f.condition = GREATER_THAN_ZERO;
							break;
						case 1:
							f.condition = LESS_THAN_ZERO;
							break;
						case 2:
							f.condition = EQUAL_ZERO;
							break;
					}
				}
			}
		}
		FormatInfo[] ret = new FormatInfo[list.size()];
		return (FormatInfo[])list.toArray(ret);
	}
	
	private static Condition createCondition(String key) {
		if (key.length() < 2) {
			return null;
		}
		int numPos = 1;
		char c1 = key.charAt(0);
		char c2 = key.charAt(1);
		if ((c1 == '>' || c1 == '<') && (c2 == '=' || c2 == '>')) {
			numPos = 2;
		}
		try {
			double d = Double.parseDouble(key.substring(numPos));
			switch (c1) {
				case '>':
					return numPos == 2 && c2 == '=' ? new GreaterEqual(d) : new GreaterThan(d);
				case '<':
					return numPos == 2 && c2 == '=' ? new LessEqual(d) : 
						numPos == 2 && c2 == '>' ? new NotEqual(d) : new LessThan(d);
				case '=':
					return new Equal(d);
				default:
					throw new IllegalStateException();
			}
		} catch (NumberFormatException e) {
			return null;
		}
	}
	
	private String formatStr;
	private Locale locale;
	private Color color;
	private Condition condition;
	private Format format;
	
	public FormatInfo(String formatStr, Locale locale, Color color, Condition condition) {
		this.formatStr = formatStr;
		this.locale = locale;
		this.color = color;
		this.condition = condition;
	}
	
	public String getFormatStr() { return this.formatStr;}
	public Locale getLocale() { return this.locale;}
	public Color getColor() { return this.color;}
	public Condition getCondition() { return this.condition;}
	
	public Format getFormat() { return this.format;}
	public void setFormat(Format f) { this.format = f;}
	
	public static abstract class Condition {
		
		protected double value;
		
		protected Condition(double d) {
			this.value = d;
		}
		
		public abstract boolean match(double d);
	}
	
	public static class Equal extends Condition {
		
		public Equal(double d) {
			super(d);
		}
		
		@Override
		public boolean match(double d) {
			return d == this.value;
		}
	}
	
	public static class NotEqual extends Condition {
		
		public NotEqual(double d) {
			super(d);
		}
		
		@Override
		public boolean match(double d) {
			return d != this.value;
		}
	}
	
	public static class LessThan extends Condition {
		
		public LessThan(double d) {
			super(d);
		}
		
		@Override
		public boolean match(double d) {
			return d < this.value;
		}
	}
	
	public static class LessEqual extends Condition {
		
		public LessEqual(double d) {
			super(d);
		}
		
		@Override
		public boolean match(double d) {
			return d <= this.value;
		}
	}
	
	public static class GreaterThan extends Condition {
		
		public GreaterThan(double d) {
			super(d);
		}
		
		@Override
		public boolean match(double d) {
			return d > this.value;
		}
	}
	
	public static class GreaterEqual extends Condition {
		
		public GreaterEqual(double d) {
			super(d);
		}
		
		@Override
		public boolean match(double d) {
			return d >= this.value;
		}
	}
	
	public static final Condition GREATER_THAN_ZERO = new GreaterThan(0);
	public static final Condition LESS_THAN_ZERO    = new LessThan(0);
	public static final Condition EQUAL_ZERO        = new Equal(0);
}
