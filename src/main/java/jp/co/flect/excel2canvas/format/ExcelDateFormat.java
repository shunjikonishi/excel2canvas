package jp.co.flect.excel2canvas.format;

import java.text.Format;
import java.text.FieldPosition;
import java.text.ParsePosition;
import java.util.List;
import java.util.ArrayList;
import java.util.Locale;
import java.util.Calendar;
import java.util.Date;

/**
 * Japanese specific date format.
 * @See http://support.microsoft.com/kb/883199/ja
 * yy	西暦の下 2 桁を表示します。
 * yyyy	西暦を 4 桁で表示します。
 * e	年を年号を元に表示します。
 * ee	年を年号を元に 2 桁の数値で表示します。
 * g	元号をアルファベットの頭文字 (M、T、S、H) で表示します。
 * gg	元号を漢字の頭文字 (明、大、昭、平) で表示します。
 * ggg	元号を漢字 (明治、大正、昭和、平成) で表示します
 * m	月を表示します。
 * mm	1 桁の月には 0 をつけて 2 桁で表示します。
 * mmm	英語の月の頭文字 3 文字 (Jan～Dec) を表示します。
 * mmmm	英語の月 (January～December) を表示します。
 * mmmmm英語の月の頭文字 (J～D) で表示します。
 * d	日にちを表示します。
 * dd   1 桁の日にちには 0 をつけて 2 桁で表示します。
 * ddd	英語の曜日の頭文字から 3 文字 (Sun～Sat) を表示します。
 * dddd	英語の曜日 (Sunday～Saturday) を表示します。
 * aaa	漢字で曜日の頭文字 (日～土) を表示します。
 * aaaa 漢字で曜日 (日曜日～土曜日) を表示します。
 * h	時刻 (0～23) を表示します。
 * hh	1 桁の時刻には 0 を付けて時刻 (00～23) を表示します。
 * m	分 (0～59) を表示します。
 * mm	1 桁の分は、0 を付けて分 (00～59) を表示します。
 * s	秒 (0～59) を表示します。
 * ss   1 桁の秒は 0 を付けて秒 (00～59) を表示します。
 * 
 * AM/PM (午前/午後)
 * AM/PM 、am/pm 、A/P 、a/p を時刻の書式記号に含めると、時刻は 12 時間表示で表示されます。また、大文字、小文字も入力したとおりに表示されます。
 * 書式記号	説明
 * h AM/PM	12 時間表示で時刻の後に AM または PM を表示します。
 * h:mm AM/PM	12 時間表示で時間と分の後に AM または PM を表示します。
 * h:mm:ss A/P	12 時間表示で時間と分と秒の後に A または P を表示します。
 * [h]:mm	24 時間を超える時間の合計を表示します。
 * [mm]:ss	60 分を超える分の合計を表示します。
 * [ss]	60 秒を超える秒の合計を表示します。
 */
public class ExcelDateFormat extends Format {
	
	private Locale locale;
	private List<Builder> list = new ArrayList<Builder>();
	
	private boolean ampm = false;
	
	public ExcelDateFormat(String str) {
		this(str, Locale.getDefault());
	}
	
	public ExcelDateFormat(String str, boolean japanese) {
		this(str, japanese ? new Locale("ja", "JP", "JP") : Locale.getDefault());
	}
	
	public ExcelDateFormat(String str, Locale l) {
		if ("ja".equals(l.getLanguage()) && (!"JP".equals(l.getCountry()) || !"JP".equals(l.getVariant()))) {
			l = new Locale("ja", "JP", "JP");
		}
		this.locale = l;
		build(str);
	}
	
	public StringBuffer format(Object obj, StringBuffer buf, FieldPosition pos) {
		Date d = (Date)obj;
		Calendar cal = Calendar.getInstance();
		cal.setTime(d);
		StringBuilder temp = new StringBuilder();
		for (Builder builder : this.list) {
			builder.append(cal, temp);
		}
		buf.append(temp);
		return buf;
	}
	
	public Object parseObject(String source, ParsePosition pos) {
		throw new UnsupportedOperationException();
	}
	
	private boolean isJapanese() {
		return "ja".equals(this.locale.getLanguage());
	}
	
	private void build(String str) {
		Builder prev = null;
		int len = str.length();
		for (int i=0; i<len; i++) {
			char c = str.charAt(i);
			switch (c) {
				case 'y':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					cnt = cnt <= 2 ? 2 : 4;
					prev = new Year(cnt);
					this.list.add(prev);
					break;
				}
				case 'g':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					if (cnt > 3) {
						cnt = 3;
					}
					prev = isJapanese() ? new JapaneseEra(cnt) : new GregorianEra();
					this.list.add(prev);
					break;
				}
				case 'e':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					prev = isJapanese() ? new JapaneseYear(cnt >= 2) : new Year(4);
					this.list.add(prev);
					break;
				}
				case 'm':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					if (cnt > 5) {
						cnt = 5;
					}
					prev = (prev instanceof Hour && cnt <= 2) ? new Minute(cnt == 2) : new Month(cnt);
					this.list.add(prev);
					break;
				}
				case 'd':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					prev = (cnt <= 2) ? new Day(cnt == 2) : new Week(false, cnt == 3);
					this.list.add(prev);
					break;
				}
				case 'a':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					if (cnt >= 3) {
						prev = new Week(true, cnt == 3);
						this.list.add(prev);
					} else if (cnt == 1 && i+1 < len && (str.charAt(i+1) == '/' || str.charAt(i+1) == 'm')) {
						i = buildAmPm(str, i);
					} else {
						this.list.add(new Const(c, cnt));
					}
					break;
				}
				case 'A':
					i = buildAmPm(str, i);
					break;
				case 'h':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					prev = new Hour(cnt >= 2);
					this.list.add(prev);
					break;
				}
				case 's':
				{
					int cnt = 1;
					while (i+1 < len && str.charAt(i+1) == c) {
						i++;
						cnt++;
					}
					if (prev instanceof Month) {
						replaceMonth((Month)prev);
					}
					prev = new Second(cnt >= 2);
					this.list.add(prev);
					break;
				}
				case '[':
				{
					if (i+1 < len) {
						int spos = i;
						char c2 = str.charAt(i+1);
						if (c2 == 'h' || c2 == 'm' || c2 == 's') {
							i++;
							int cnt = 1;
							while (i+1 < len && str.charAt(i+1) == c2) {
								i++;
								cnt++;
							}
							if (i == len || str.charAt(i) != ']') {
								this.list.add(new Const(str.substring(spos, i)));
							} else {
								i++;
								prev = c2 == 'h' ? new SpecialHour(cnt >= 2) :
									c2 == 'm' ? new SpecialMinute(cnt >= 2) : new SpecialSecond(cnt >= 2);
								this.list.add(prev);
							}
						} else {
							int epos = str.indexOf(']', i);
							if (epos == -1) {
								epos = str.length();
							}
							i = epos;
						}
					} else{
						this.list.add(new Const(c, 1));
					}
					break;
				}
				case '.':
				{
					if ((prev instanceof Second || prev instanceof SpecialSecond) && i+1 < len && str.charAt(i+1) == '0') {
						i++;
						int cnt = 1;
						while (i+1 < len && str.charAt(i+1) == '0') {
							i++;
							cnt++;
						}
						prev = new MilliSecond(cnt);
						this.list.add(prev);
					} else {
						this.list.add(new Const(c, 1));
					}
					break;
				}
				case '\\':
				{
					if (i+1 < len) {
						this.list.add(new Const(str.charAt(i+1), 1));
						i++;
					}
					break;
				}
				case '"':
				{
					if (i+1 <len) {
						int idx = str.indexOf('"', i+1);
						if (idx == -1) {
							idx = len;
						}
						this.list.add(new Const(str.substring(i+1, idx)));
						i = idx;
					}
					break;
				}
				default:
					this.list.add(new Const(c, 1));
					break;
			}
		}
	}
	
	private void replaceMonth(Month m) {
		int idx = this.list.lastIndexOf(m);
		this.list.set(idx, m.toMinute());
	}
	
	private int buildAmPm(String str, int idx) {
		String am = null;
		StringBuilder buf = new StringBuilder();
		for (int i=idx; i<str.length(); i++) {
			char c = str.charAt(i);
			switch (c) {
				case 'a':
				case 'A':
					if (am == null && buf.length() == 0) {
						buf.append(c);
					} else {
						return idx;
					}
					break;
				case 'p':
				case 'P':
					if (am != null && buf.length() == 0) {
						buf.append(c);
					} else {
						return idx;
					}
					break;
				case 'm':
				case 'M':
					if (buf.length() == 1) {
						buf.append(c);
					} else {
						return idx;
					}
					break;
				case '/':
					if (am == null && buf.length() > 0) {
						am = buf.toString();
						buf.setLength(0);
					} else {
						return idx;
					}
					break;
				default:
					if (am != null && buf.length() > 0) {
						this.list.add(new AmPm(am, buf.toString()));
						this.ampm = true;
						return i;
					} else {
						return idx;
					}
			}
		}
		if (am != null && buf.length() > 0) {
			this.list.add(new AmPm(am, buf.toString()));
			this.ampm = true;
			return str.length()-1;
		} else {
			return idx;
		}
	}
	public interface Builder {
		
		public void append(Calendar cal, StringBuilder buf);
		
	}
	
	private static class Const implements Builder {
		
		private String str;
		
		public Const(char c, int cnt) {
			StringBuilder buf = new StringBuilder();
			for (int i=0; i<cnt; i++) {
				buf.append(c);
			}
			this.str = buf.toString();
		}
		
		public Const(String str) {
			this.str = str;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			buf.append(str);
		}
	}
	
	private static class Year implements Builder {
		
		private int len;
		
		public Year(int len) {
			this.len = len;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.YEAR);
			if (this.len == 2) {
				n = n % 100;
				if (n < 10) {
					buf.append("0");
				}
			}
			buf.append(Integer.toString(n));
		}
	}
	
	private static class GregorianEra implements Builder {
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.ERA);
			buf.append(n == 1 ? "AD" : "BC");
		}
	}
	
	private static final String[] SHORT_ERA = { "M", "T", "S", "H" };
	private static final String[] LONG_ERA = { "明治", "大正", "昭和", "平成" };
	
	private class JapaneseEra implements Builder {
		
		private int len;
		
		public JapaneseEra(int len) {
			this.len = len;
		}
			
		public void append(Calendar gCal, StringBuilder buf) {
			Calendar cal = Calendar.getInstance(ExcelDateFormat.this.locale);
			cal.setTime(gCal.getTime());
			int n = cal.get(Calendar.ERA);
			String era = null;
			if (n >= 1 && n <= 4) {
				switch (this.len) {
					case 1:
						era = SHORT_ERA[n-1];
						break;
					case 2:
						era = LONG_ERA[n-1].substring(0, 1);
						break;
					case 3:
						era = LONG_ERA[n-1];
						break;
				}
			}
			if (era != null) {
				buf.append(era);
			}
		}
	}
	
	private class JapaneseYear implements Builder {
		
		private boolean forceTwo;
		
		public JapaneseYear(boolean forceTwo) {
			this.forceTwo = forceTwo;
		}
		
		public void append(Calendar gCal, StringBuilder buf) {
			Calendar cal = Calendar.getInstance(ExcelDateFormat.this.locale);
			cal.setTime(gCal.getTime());
			int n = cal.get(Calendar.YEAR);
			if (forceTwo && n < 10) {
				buf.append("0");
			}
			buf.append(Integer.toString(n));
		}
	}
	
	private static final String[] MONTH = {
		"January", "February", "March", "April", "May", "June",
		"July", "August", "September", "October", "November", "December"
	};
	
	private static class Month implements Builder {
		
		private int len;
		
		public Month(int len) {
			this.len = len;
		}
		
		public Minute toMinute() {
			return new Minute(this.len >= 2);
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.MONTH);
			if (this.len <= 2) {
				if (this.len == 2 && n+1 < 10) {
					buf.append("0");
				}
				buf.append(Integer.toString(n+1));
			} else {
				String str = MONTH[n];
				switch (this.len) {
					case 3:
						str = str.substring(0, 3);
						break;
					case 5:
						str = str.substring(0, 1);
						break;
				}
				buf.append(str);
			}
		}
		
	}
	
	private static class Day implements Builder {
		
		private boolean forceTwo;
		
		public Day(boolean forceTwo) {
			this.forceTwo = forceTwo;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.DAY_OF_MONTH);
			if (forceTwo && n < 10) {
				buf.append("0");
			}
			buf.append(Integer.toString(n));
		}
		
	}
	
	private static final String[] ENGLISH_WEEK = {
		"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
	};
	
	private static final String[] JAPANESE_WEEK = {
		"日", "月", "火", "水", "木", "金", "土"
	};
	
	private static class Week implements Builder {
		
		private boolean japanese;
		private boolean bShort;
		
		public Week(boolean japanese, boolean bShort) {
			this.japanese = japanese;
			this.bShort = bShort;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.DAY_OF_WEEK) - 1;
			String s = this.japanese ? JAPANESE_WEEK[n] : ENGLISH_WEEK[n];
			if (!japanese && bShort) {
				s = s.substring(0, 3);
			}
			buf.append(s);
			if (japanese && !bShort) {
				buf.append("曜日");
			}
		}
	}
	
	private class Hour implements Builder {
		
		private boolean forceTwo;
		
		public Hour(boolean forceTwo) {
			this.forceTwo = forceTwo;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.HOUR_OF_DAY);
			if (ExcelDateFormat.this.ampm && n > 12) {
				n -= 12;
			}
			if (forceTwo && n < 10) {
				buf.append("0");
			}
			buf.append(Integer.toString(n));
		}
		
	}
	
	private static class Minute implements Builder {
		
		private boolean forceTwo;
		
		public Minute(boolean forceTwo) {
			this.forceTwo = forceTwo;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.MINUTE);
			if (forceTwo && n < 10) {
				buf.append("0");
			}
			buf.append(Integer.toString(n));
		}
		
	}
	
	private static class Second implements Builder {
		
		private boolean forceTwo;
		
		public Second(boolean forceTwo) {
			this.forceTwo = forceTwo;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.SECOND);
			if (forceTwo && n < 10) {
				buf.append("0");
			}
			buf.append(Integer.toString(n));
		}
		
	}
	
	private static class SpecialHour implements Builder {
		
		private boolean forceTwo;
		
		public SpecialHour(boolean forceTwo) {
			this.forceTwo = forceTwo;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = calc(cal);
			if (forceTwo && n < 10) {
				buf.append("0");
			}
			buf.append(Integer.toString(n));
		}
		
		protected int calc(Calendar cal) {
			int d = cal.get(Calendar.DAY_OF_YEAR);
			int h = cal.get(Calendar.HOUR_OF_DAY);
			return d * 24 + h;
		}
	}
	
	private static class SpecialMinute extends SpecialHour {
		
		public SpecialMinute(boolean forceTwo) {
			super(forceTwo);
		}
		
		protected int calc(Calendar cal) {
			int h = super.calc(cal);
			int m = cal.get(Calendar.MINUTE);
			return h * 60 + m;
		}
	}
	
	private static class SpecialSecond extends SpecialMinute {
		
		public SpecialSecond(boolean forceTwo) {
			super(forceTwo);
		}
		
		protected int calc(Calendar cal) {
			int m = super.calc(cal);
			int s = cal.get(Calendar.SECOND);
			return m * 60 + s;
		}
	}
	
	private static class AmPm implements Builder {
		
		private String am;
		private String pm;
		
		public AmPm(String am, String pm) {
			this.am = am;
			this.pm = pm;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.HOUR_OF_DAY);
			String s = n < 12 ? am : pm;
			buf.append(s);
		}
		
	}
	
	private static class MilliSecond implements Builder {
		
		private int len;
		
		public MilliSecond(int len) {
			this.len = len;
		}
		
		public void append(Calendar cal, StringBuilder buf) {
			int n = cal.get(Calendar.MILLISECOND);
			String s = Integer.toString(n);
			if (this.len < s.length()) {
				s = s.substring(0, this.len);
			}
			buf.append(".").append(s);
		}
	}
}
