package jp.co.flect.excel2canvas;

import java.util.Date;
import java.util.Locale;
import java.util.Map;
import java.util.HashMap;
import java.text.SimpleDateFormat;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.formula.eval.NotImplementedException;

/**
 * Helper class for get cell values.
 */
public class CellValueHelper {
	
	private static final String[] DATE_FORMATS = {
		"yyyy-MM-dd",
		"yyyy-MM-dd HH:mm:ss",
		"yyyy/MM/dd",
		"yyyy/MM/dd HH:mm:ss",
		"HH:mm:ss"
	};
	
	public static boolean isTextFormat(Cell cell) {
		if (cell == null) return false;

		CellStyle style = cell.getCellStyle();
		if (style == null) return false;

		return style.getDataFormat() == 49 && "@".equals(style.getDataFormatString());
	}

	private Workbook workbook;
	private FormulaEvaluator evaluator;
	private DataFormatterEx dataFormatter;
	private HashMap<String, FormattedValue> cached = null;
	private SimpleDateFormat[] dateFormats = null;
	private ExceptionHandler exHandler = null;
	
	public CellValueHelper(Workbook workbook, boolean cache) {
		this(workbook, cache, Locale.getDefault());
	}
	
	public CellValueHelper(Workbook workbook, boolean cache, Locale locale) {
		this.workbook = workbook;
		this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		this.dataFormatter = new DataFormatterEx(locale);
		if (cache) {
			this.cached = new HashMap<String, FormattedValue>();
		}
	}
	
	private SimpleDateFormat getDateFormat(int idx) {
		if (this.dateFormats == null) {
			this.dateFormats = new SimpleDateFormat[DATE_FORMATS.length];
		}
		SimpleDateFormat ret = this.dateFormats[idx];
		if (ret == null) {
			ret = new SimpleDateFormat(DATE_FORMATS[idx]);
			this.dateFormats[idx] = ret;
		}
		return ret;
	}
	
	public void clearCache() {
		if (this.cached != null) {
			this.cached.clear();
		}
		//FormulaEvaluator keeps calculated values.
		this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	}

	public FormulaEvaluator getEvaluator() { return this.evaluator;}
	
	public FormattedValue getFormattedValue(Cell cell) {
		return getFormattedValue(cell, cell.getCellType());
	}
	
	private FormattedValue getFormattedValue(Cell cell, int type) {
		try {
			switch(type) {
				case Cell.CELL_TYPE_BLANK:
				case Cell.CELL_TYPE_ERROR:
				case Cell.CELL_TYPE_STRING:
				case Cell.CELL_TYPE_BOOLEAN:
				case Cell.CELL_TYPE_NUMERIC:
					return dataFormatter.formatCellValue(cell);
				case Cell.CELL_TYPE_FORMULA:
					FormattedValue ret = null;
					String key = null;
					try {
						if (this.cached != null) {
							key = cell.getSheet().getSheetName() + "!" + ExcelUtils.pointToName(cell.getColumnIndex(), cell.getRowIndex());
							ret = this.cached.get(key);
							if (ret != null) {
								return ret;
							}
						}
						ret = dataFormatter.formatCellValue(cell, evaluator);
					} catch (NotImplementedException e) {
						handleException(e);
					} catch (Exception e) {
						handleException(e);
					}
					if (ret == null) {
						int cachedType = cell.getCachedFormulaResultType();
						if (cachedType == Cell.CELL_TYPE_FORMULA) {
							ret = new FormattedValue(cell.getCellFormula(), FormattedValue.Type.FORMULA, cell.getCellFormula());
						} else {
							ret = getFormattedValue(cell, cachedType);
						}
						ret.setFormula(cell.getCellFormula());
					}
					if (this.cached != null) {
						this.cached.put(key, ret);
					}
					return ret;
				default:
					break;
			}
		} catch (Exception e) {
			handleException(e);
			return new FormattedValue(e.toString(), FormattedValue.Type.ERROR, e.toString());
		}
		throw new IllegalStateException();
	}
	
	public void setString(Cell cell, String value) {
		if (value == null || value.length() == 0) {
			cell.setCellValue("");
			return;
		}
		char firstChar = value.charAt(0);
		if (!isTextFormat(cell) && (firstChar == '-' || firstChar == '.' || (firstChar >= '0' && firstChar <= '9'))) {
			try {
				cell.setCellValue(Double.parseDouble(value));
				return;
			} catch (Exception e) {
				//ignore
			}
			if (firstChar >= '0' && firstChar <= '9') {
				for (int i=0; i<DATE_FORMATS.length; i++) {
					SimpleDateFormat fmt = getDateFormat(i);
					try {
						Date d = fmt.parse(value);
						cell.setCellValue(d);
						return;
					} catch (Exception e) {
						//ignore
					}
				}
			}
		} else if (firstChar == 't' && value.equals("true")) {
			cell.setCellValue(true);
			return;
		} else if (firstChar == 'f' && value.equals("false")) {
			cell.setCellValue(false);
			return;
		} else if (firstChar == '=') {
			cell.setCellFormula(value);
			return;
		} else if (firstChar == '\'') {
			cell.setCellValue(value.substring(1));
			return;
		}
		cell.setCellValue(value);
	}
	
	public boolean isEmptyCell(Cell cell) {
		if (cell == null) {
			return true;
		}
		if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			return false;
		}
		String value = getFormattedValue(cell).getValue();
		return value == null || value.length() == 0;
	}

	public boolean isEmptyCell(CellReference cRef) {
		Sheet sheet = workbook.getSheet(cRef.getSheetName());
		if (sheet == null) {
			return false;
		}
		return isEmptyCell(sheet, cRef);
	}

	public boolean isEmptyCell(Sheet sheet, CellReference cRef) {
		Row row = sheet.getRow(cRef.getRow());
		if (row == null) {
			return true;
		}
		return isEmptyCell(row.getCell(cRef.getCol()));
	}
	
	public ExceptionHandler getExceptionHandler() { return this.exHandler;}
	public void setExceptionHandler(ExceptionHandler v) { this.exHandler = v;}

	public interface ExceptionHandler {
		public void handle(Exception e);
	}

	private void handleException(Exception e) {
		if (this.exHandler != null) {
			this.exHandler.handle(e);
		} else if (e instanceof NotImplementedException) {
			int stack = 0;
			while (e.getCause() != null) {
				e = (NotImplementedException)e.getCause();
				stack++;
			}
			System.err.println("!!! Unsupported Formula !!! - " + e.getMessage() + " - " + stack);
		} else {
			System.err.println("!!! Unknown Error !!!");
			e.printStackTrace();
		}
	}
}
