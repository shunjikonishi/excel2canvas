package jp.co.flect.excel2canvas.format;

/**
 * FormatHolder
 */
public class FormatHolder {
	
	private FormatInfo[] formats;
	
	public FormatHolder(FormatInfo[] formats) {
		this.formats = formats;
	}
	
	public FormatInfo getFormatInfo() {
		return this.formats[0];
	}
	
	public FormatInfo getFormatInfo(double d) {
		for (FormatInfo f : this.formats) {
			FormatInfo.Condition cond = f.getCondition();
			if (cond == null || cond.match(d)) {
				return f;
			}
		}
		return this.formats[0];
	}
	
}
