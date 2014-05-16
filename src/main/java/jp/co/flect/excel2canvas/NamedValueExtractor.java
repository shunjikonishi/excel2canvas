package jp.co.flect.excel2canvas;

import java.io.File;
import java.io.InputStream;
import java.io.IOException;
import java.util.Locale;
import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.util.Date;
import java.text.SimpleDateFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import com.google.gson.Gson;

public class NamedValueExtractor {

	private Locale locale;
	private boolean convertString = false;
	private boolean includeFormulaValue = false;

	public NamedValueExtractor() {
		this(Locale.getDefault());
	}

	public NamedValueExtractor(Locale locale) {
		this.locale = locale;
	}

	public Locale getLocale() { return this.locale;}
	public void setLocale(Locale l) { this.locale = l;}

	public boolean isConvertString() { return this.convertString;}
	public void setConvertString(boolean b) { this.convertString = b;}

	public boolean isIncludeFormulaValue() { return this.includeFormulaValue;}
	public void setIncludeFormulaValue(boolean b) { this.includeFormulaValue = b;}

	public Map<String, Object[]> extract(Workbook workbook) {
		CellValueHelper helper = new CellValueHelper(workbook, true, this.locale);
		List<NamedCellInfo> list = ExcelUtils.createNamedCellList(workbook);
		Map<String, Object[]> map = new HashMap<String, Object[]>();
		for (NamedCellInfo info : list) {
			int currentLine = -1;
			int includeSize = 0;
			boolean hasValue = false;
			Sheet sheet = workbook.getSheet(info.
				getSheetName());
			List<Object> values = new ArrayList<Object>();
			for (CellReference cRef : info.getCellList()) {
				Cell cell = ExcelUtils.getOrCreateCell(sheet, cRef.getRow(), cRef.getCol());
				if (cRef.getRow() != currentLine) {
					currentLine = cRef.getRow();
					if (hasValue) {
						includeSize = values.size();
					}
					hasValue = false;
				}
				if (cell.getCellType() == Cell.CELL_TYPE_FORMULA && info.getCellList().size() > 1 && !includeFormulaValue) {
					continue;
				}
				FormattedValue fv = helper.getFormattedValue(cell);
				Object value = null;
				if (this.convertString) {
					value = fv.getValue();
				} else if (fv.getRawData() != null) {
					value = fv.getRawData();
				} else {
					switch (fv.getType()) {
						case NUMBER:
							value = cell.getNumericCellValue();
							break;
						case DATE:
							value = cell.getDateCellValue();
							break;
						case BOOLEAN:
							value = cell.getBooleanCellValue();
							break;
						default:
							value = fv.getValue();
							break;
					}
				}
				if (value != null && value.toString().length() > 0) {
					hasValue = true;
				}
				values.add(value);
			}
			if (hasValue) {
				includeSize = values.size();
			}
			map.put(info.getName(), values.subList(0, includeSize).toArray());
		}
		return map;
	}

	public Map<String, Object[]> extract(File f) throws IOException, InvalidFormatException {
		return extract(ExcelUtils.createWorkbook(f));
	}

	public Map<String, Object[]> extract(InputStream is) throws IOException, InvalidFormatException {
		return extract(ExcelUtils.createWorkbook(is));
	}

	public static String toJson(Map<String, Object[]> map) {
		SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.sss");
		Map<String, Object> ret = new HashMap<String, Object>();
		for (Map.Entry<String, Object[]> entry : map.entrySet()) {
			String key = entry.getKey();
			Object[] values = entry.getValue();
			if (values.length == 0) {
				ret.put(key, null);
			} else if (values.length == 1) {
				ret.put(key, convert(values[0], format));
			} else {
				Object[] newValues = new Object[values.length];
				for (int i=0; i<values.length; i++) {
					newValues[i] = convert(values[i], format);
				}
				ret.put(key, newValues);
			}
		}
		return new Gson().toJson(ret);
	}

	private static Object convert(Object o, SimpleDateFormat format) {
		if (o instanceof Date) {
			return format.format((Date)o);
		} else if (o instanceof String && o.toString().length() == 0) {
			return null;
		}
		return o;
	}
}