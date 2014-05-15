package jp.co.flect.excel2canvas;

import java.io.File;
import java.io.InputStream;
import java.io.IOException;
import java.util.Locale;
import java.util.Map;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

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
			Sheet sheet = workbook.getSheet(info.getSheetName());
			List<Object> values = new ArrayList<Object>();
			for (CellReference cRef : info.getCellList()) {
				Cell cell = ExcelUtils.getOrCreateCell(sheet, cRef.getRow(), cRef.getCol());
				if (cell.getCellType() == Cell.CELL_TYPE_FORMULA && !includeFormulaValue) {
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
				values.add(value);
			}
			map.put(info.getName(), values.toArray());
		}
		return map;
	}

	public Map<String, Object[]> extract(File f) throws IOException, InvalidFormatException {
		return extract(ExcelUtils.createWorkbook(f));
	}

	public Map<String, Object[]> extract(InputStream is) throws IOException, InvalidFormatException {
		return extract(ExcelUtils.createWorkbook(is));
	}
}