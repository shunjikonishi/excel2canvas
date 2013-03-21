package jp.co.flect.excel2canvas;
	
import java.text.Format;
import java.text.SimpleDateFormat;
import java.text.FieldPosition;
import java.text.ParsePosition;
import java.util.Locale;
import java.util.Date;
import java.util.Map;
import java.util.HashMap;
import java.awt.Color;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import jp.co.flect.excel2canvas.format.ExcelDateFormat;
import jp.co.flect.excel2canvas.format.FormatInfo;
import jp.co.flect.excel2canvas.format.FormatHolder;

/**
 * Japanese specific date formatter.
 */
class DataFormatterEx{
	
	private static final boolean DEBUG;
	
	static {
		DEBUG = "true".equals(System.getProperty(DataFormatterEx.class.getName() + ".DEBUG"));
	}
	
	private DataFormatter numberFormatter;
	private Map<String, FormatHolder> formats = new HashMap<String, FormatHolder>();
	private Locale locale;
	
	public DataFormatterEx() {
		this(Locale.getDefault(), false);
	}
	
	public DataFormatterEx(boolean emulateCsv) {
		this(Locale.getDefault(), emulateCsv);
	}
	
	public DataFormatterEx(Locale locale) {
		this(locale, false);
	}
	
	public DataFormatterEx(Locale locale, boolean emulateCsv) {
		//和暦対応
		if ("ja".equals(locale.getLanguage())) {
			if (!"JP".equals(locale.getCountry()) || !"JP".equals(locale.getVariant())) {
				locale = new Locale("ja", "JP", "JP");
			}
		}
		this.locale = locale;
		this.numberFormatter = new DataFormatter(locale, emulateCsv);
	}
	
	/**
	 * DataFormatには各国で予約されていて書式文字列が返ってこないものがあるそれぞれの国でどのような書式が登録されているかは不明
	 * とりあえず日本語にのみ対応する
	 */
	private String getLocalizedDateFormat(short idx) {
		if ("ja".equals(this.locale.getLanguage())) {
			switch (idx) {
				case 30: return "m/d/yy";
				case 31: return "yyyy\"年\"m\"月\"d\"日\"";
				case 32: return "h\"時\"mm\"分\"";
				case 33: return "h\"時\"mm\"分\"ss\"秒\"";
				case 55: return "yyyy\"年\"m\"月\"";
				case 56: return "m\"月\"d\"日\"";
				case 57: return "[$-411]ge.m.d";
				case 58: return "[$-411]ggge\"年\"m\"月\"d\"日\"";
			}
		}
		return null;
	}
	
    public FormattedValue formatCellValue(Cell cell) {
		return formatCellValue(cell, null);
	}
	
    public FormattedValue formatCellValue(Cell cell, FormulaEvaluator evaluator) {
		FormattedValue ret = doFormatCellValue(cell, evaluator);
		if (cell != null && cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			ret.setFormula(cell.getCellFormula());
		}
		if (DEBUG && cell != null && cell.getCellStyle() != null) {
			int cellType = cell.getCellType();
			if (cellType == Cell.CELL_TYPE_FORMULA) {
				cellType = cell.getCachedFormulaResultType();
			}
			if (cellType == Cell.CELL_TYPE_NUMERIC) {
				System.out.println("Format: " + 
					cell.getCellStyle().getDataFormat() + ": " + 
					cell.getCellStyle().getDataFormatString() + " --- " + 
					ExcelUtils.isCellDateFormatted(cell) + ":" + 
					DateUtil.isCellDateFormatted(cell) + ":" + 
					DateUtil.isInternalDateFormat(cell.getCellStyle().getDataFormat()) + " --- " + 
					ret); 
			}
		}
		return ret;
	}
	
	private FormattedValue doFormatCellValue(Cell cell, FormulaEvaluator evaluator) {
		if (cell == null) {
			return FormattedValue.EMPTY;
		}
		int cellType = cell.getCellType();
		if (cellType == Cell.CELL_TYPE_FORMULA) {
			if (evaluator == null) {
				cellType = cell.getCachedFormulaResultType();
				if (cellType == Cell.CELL_TYPE_FORMULA) {
					return new FormattedValue(cell.getCellFormula(), FormattedValue.Type.FORMULA, cell.getCellFormula());
				}
			} else {
				cellType = evaluator.evaluateFormulaCell(cell);
			}
		}
		String ret = "";
		Object rawdata = null;
		FormattedValue.Type type = null;
		Color color = null;
		if (cellType == Cell.CELL_TYPE_NUMERIC) {
			if (cell.getCellStyle() == null) {
				ret = this.numberFormatter.formatCellValue(cell, evaluator);
				type = FormattedValue.Type.NUMBER;
			} else {
				boolean bDate = false;
				short idx = cell.getCellStyle().getDataFormat();
				String fmt = getLocalizedDateFormat(idx);
				//InternalDateFormatはPOIにまかせる
				bDate = DateUtil.isValidExcelDate(cell.getNumericCellValue()) &&
					(fmt != null || (idx > 0x31 && ExcelUtils.isADateFormat(idx, cell.getCellStyle().getDataFormatString())));
				
				if (bDate) {
					ret = formatDate(cell, fmt);
					rawdata = cell.getDateCellValue();
					type = FormattedValue.Type.DATE;
				} else {
					ret = this.numberFormatter.formatCellValue(cell, evaluator);
					rawdata = cell.getNumericCellValue();
					type = FormattedValue.Type.NUMBER;
				}
				FormatHolder holder = getFormatHolder(cell, fmt);
				if (holder != null) {
					color = holder.getFormatInfo(cell.getNumericCellValue()).getColor();
				}
			}
		} else {
			switch (cellType) {
				case Cell.CELL_TYPE_BLANK:
					return FormattedValue.EMPTY;
				case Cell.CELL_TYPE_ERROR:
					ret = String.valueOf(cell.getErrorCellValue());
					type = FormattedValue.Type.ERROR;
					break;
				case Cell.CELL_TYPE_STRING:
					ret = cell.getRichStringCellValue().getString();
					type = FormattedValue.Type.STRING;
					break;
				case Cell.CELL_TYPE_BOOLEAN:
					ret = String.valueOf(cell.getBooleanCellValue());
					type = FormattedValue.Type.BOOLEAN;
					break;
				default:
					throw new IllegalStateException();
			}
		}
		return new FormattedValue(ret, type, rawdata != null ? rawdata : ret, color);
	}
	
	private FormatHolder getFormatHolder(Cell cell, String formatStr) {
		if (formatStr == null) {
			formatStr = cell.getCellStyle().getDataFormatString();
		}
		if (formatStr == null) {
			return null;
		}
		FormatHolder holder = formats.get(formatStr);
		if (holder == null) {
			FormatInfo[] infos = FormatInfo.parse(cell.getSheet().getWorkbook(), formatStr);
			infos[0].setFormat(new ExcelDateFormat(infos[0].getFormatStr(), infos[0].getLocale() != null ? infos[0].getLocale() : this.locale));
			holder = new FormatHolder(infos);
			this.formats.put(formatStr, holder);
		}
		return holder;
	}
	
	private String formatDate(Cell cell, String formatStr) {
		Date d = cell.getDateCellValue();
		FormatHolder holder = getFormatHolder(cell, formatStr);
		if (holder == null) {
			return d.toString();
		} else {
			return holder.getFormatInfo().getFormat().format(d);
		}
	}
	
}