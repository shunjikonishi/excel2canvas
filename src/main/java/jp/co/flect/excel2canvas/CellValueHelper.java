package jp.co.flect.excel2canvas;

import java.util.Locale;
import java.util.Map;
import java.util.HashMap;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.formula.eval.NotImplementedException;

public class CellValueHelper {
	
	private FormulaEvaluator evaluator;
	private DataFormatterEx dataFormatter;
	private HashMap<String, FormattedValue> cached = null;
	
	public CellValueHelper(Workbook workbook, boolean cache) {
		this(workbook, cache, Locale.getDefault());
	}
	
	public CellValueHelper(Workbook workbook, boolean cache, Locale locale) {
		this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		this.dataFormatter = new DataFormatterEx(locale);
		if (cache) {
			this.cached = new HashMap<String, FormattedValue>();
		}
	}
	
	public void clearCache() {
		if (this.cached != null) {
			this.cached.clear();
		}
	}
	
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
						int stack = 0;
						while (e.getCause() != null) {
							e = (NotImplementedException)e.getCause();
							stack++;
						}
						System.out.println("!!! Unsupported Formula !!! - " + e.getMessage() + " - " + stack);
					} catch (Exception e) {
						System.err.println("!!! Unknown Error !!!");
						e.printStackTrace();
					}
					if (ret == null) {
						int cachedType = cell.getCachedFormulaResultType();
						if (cachedType == Cell.CELL_TYPE_FORMULA) {
							ret = new FormattedValue(cell.getCellFormula(), FormattedValue.Type.FORMULA, cell.getCellFormula());
						} else {
							ret = getFormattedValue(cell, cachedType);
						}
					}
					if (this.cached != null) {
						this.cached.put(key, ret);
					}
					return ret;
				default:
					break;
			}
		} catch (Exception e) {
			System.err.println("!!! Unknown error !!!");
			e.printStackTrace();
			return new FormattedValue(e.toString(), FormattedValue.Type.ERROR, e.toString());
		}
		throw new IllegalStateException();
	}
}
