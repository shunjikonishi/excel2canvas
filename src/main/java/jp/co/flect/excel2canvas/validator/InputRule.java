package jp.co.flect.excel2canvas.validator;

import static org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType.*;
import static org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType.*;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.joda.time.DateTime;
import org.joda.time.LocalTime;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;
import java.util.ArrayList;
import com.google.gson.Gson;
import org.w3c.dom.Node;

import jp.co.flect.excel2canvas.ExcelUtils;
import jp.co.flect.excel2canvas.CellValueHelper;

public class InputRule {

	private boolean empty;
	private String errTitle;
	private String errText;
	private int errStyle;

	private String pmTitle;
	private String pmText;

	private String[] regionsStr;
	private String[] list;

	private String f1;
	private String f2;

	private int op;
	private int vt;

	private transient CellRangeAddressList regions;
	private transient Validator validator;

	public InputRule(Sheet sheet, XSSFDataValidation dv) {
		empty = dv.getEmptyCellAllowed();
		errTitle = dv.getErrorBoxTitle();
		errText = dv.getErrorBoxText();
		errStyle = dv.getErrorStyle();
		pmTitle = dv.getPromptBoxTitle();
		pmText = dv.getPromptBoxText();

		regions = dv.getRegions();
		regionsStr = new String[regions.countRanges()];
		int idx = 0;
		for (CellRangeAddress cell : regions.getCellRangeAddresses()) {
			regionsStr[idx++] = cell.formatAsString();
		}

		DataValidationConstraint vc = dv.getValidationConstraint();
		f1 = vc.getFormula1();
		f2 = vc.getFormula2();
		op = vc.getOperator();
		vt = vc.getValidationType();
		if (vt == DataValidationConstraint.ValidationType.LIST) {
			list = buildList(sheet, f1);
		}
	}

	public boolean getAllowEmpty() { return empty;}
	public String getErrorTitle() { return errTitle;}
	public String getErrorText() { return errText;}
	public int getErrorStyle() { return errStyle;}

	public String getPromptTitle() { return pmTitle;}
	public String getPromptText() { return pmText;}

	public String[] getRegions() { return regionsStr;}
	public String[] getList() { return list;}

	public String getFormula1() { return f1;}
	public String getFormula2() { return f2;}

	public int getOperator() { return op;}
	public int getValidationType() { return vt;}

	public void validator(String value) throws Exception {
		if (value == null || value.length() == 0) {
			if (empty) {
				return;
			}
			throw new Exception("Value is required.");
		}
		if (validator == null) {
			validator = createValidator();
		}
		if (!validator.validate(value)) {
			throw new Exception("Invalid value.");
		}
	}

	private Validator createValidator() {
		switch (vt) {
			case ANY:
				return new StringValidator();
			case INTEGER:
				return new IntegerValidator();
			case DECIMAL:
				return new DecimalValidator();
			case LIST:
				return new ListValidator();
			case DATE:
				return new DateValidator();
			case TIME:
				return new TimeValidator();
			case TEXT_LENGTH:
				return new TextLengthValidator();
			case FORMULA:
			default:
				return new FormulaValidator();
		}
	}

	private static interface Validator {
		public boolean validate(String value) throws Exception;
	}

	private class StringValidator implements Validator {
		public boolean validate(String value) throws Exception {
			String v = value;
			switch (op) {
				case BETWEEN:
					return v.compareTo(f1) >= 0 && v.compareTo(f2) <= 0;
				case NOT_BETWEEN:
					return v.compareTo(f1) <= 0 && v.compareTo(f2) >= 0;
				case EQUAL:
					return v.compareTo(f1) == 0;
				case NOT_EQUAL:
					return v.compareTo(f1) != 0;
				case GREATER_THAN:
					return v.compareTo(f1) > 0;
				case LESS_THAN:
					return v.compareTo(f1) < 0;
				case GREATER_OR_EQUAL:
					return v.compareTo(f1) >= 0;
				case LESS_OR_EQUAL:
					return v.compareTo(f1) <= 0;
				default:
					return true;
			}
		}
	}
	private class IntegerValidator implements Validator {
		private long l1 = f1 == null ? 0 : Long.parseLong(f1);
		private long l2 = f2 == null ? 0 : Long.parseLong(f2);

		public boolean validate(String value) throws Exception {
			long v = Long.parseLong(value);
			switch (op) {
				case BETWEEN:
					return v >= l1 && v <= l2;
				case NOT_BETWEEN:
					return v <= l1 && v >= l2;
				case EQUAL:
					return v == l1;
				case NOT_EQUAL:
					return v == l1;
				case GREATER_THAN:
					return v > l1;
				case LESS_THAN:
					return v < l1;
				case GREATER_OR_EQUAL:
					return v >= l1;
				case LESS_OR_EQUAL:
					return v <= l1;
				default:
					return true;
			}
		}
	}
	private class DecimalValidator implements Validator {
		private BigDecimal d1 = f1 == null ? null : new BigDecimal(f1);
		private BigDecimal d2 = f2 == null ? null : new BigDecimal(f2);

		public boolean validate(String value) throws Exception {
			BigDecimal v = new BigDecimal(value);
			switch (op) {
				case BETWEEN:
					return v.compareTo(d1) >= 0 && v.compareTo(d2) <= 0;
				case NOT_BETWEEN:
					return v.compareTo(d1) <= 0 && v.compareTo(d2) >= 0;
				case EQUAL:
					return v.compareTo(d1) == 0;
				case NOT_EQUAL:
					return v.compareTo(d1) != 0;
				case GREATER_THAN:
					return v.compareTo(d1) > 0;
				case LESS_THAN:
					return v.compareTo(d1) < 0;
				case GREATER_OR_EQUAL:
					return v.compareTo(d1) >= 0;
				case LESS_OR_EQUAL:
					return v.compareTo(d1) <= 0;
				default:
					return true;
			}
		}
	}
	private class ListValidator implements Validator {
		public boolean validate(String value) throws Exception {
			if (list == null || list.length == 0) {
				return false;
			}
			for (String v : list) {
				if (v.equals(value)) {
					return true;
				}
			}
			return false;
		}
	}
	private class DateValidator implements Validator {
		private DateTime d1 = f1 == null ? null : DateTime.parse(f1);
		private DateTime d2 = f2 == null ? null : DateTime.parse(f2);

		public boolean validate(String value) throws Exception {
			DateTime v = DateTime.parse(value);
			switch (op) {
				case BETWEEN:
					return v.compareTo(d1) >= 0 && v.compareTo(d2) <= 0;
				case NOT_BETWEEN:
					return v.compareTo(d1) <= 0 && v.compareTo(d2) >= 0;
				case EQUAL:
					return v.compareTo(d1) == 0;
				case NOT_EQUAL:
					return v.compareTo(d1) != 0;
				case GREATER_THAN:
					return v.compareTo(d1) > 0;
				case LESS_THAN:
					return v.compareTo(d1) < 0;
				case GREATER_OR_EQUAL:
					return v.compareTo(d1) >= 0;
				case LESS_OR_EQUAL:
					return v.compareTo(d1) <= 0;
				default:
					return true;
			}
		}
	}
	private class TimeValidator implements Validator {
		private LocalTime t1 = f1 == null ? null : LocalTime.parse(f1);
		private LocalTime t2 = f1 == null ? null : LocalTime.parse(f2);

		public boolean validate(String value) throws Exception {
			LocalTime v = LocalTime.parse(value);
			switch (op) {
				case BETWEEN:
					return v.compareTo(t1) >= 0 && v.compareTo(t2) <= 0;
				case NOT_BETWEEN:
					return v.compareTo(t1) <= 0 && v.compareTo(t2) >= 0;
				case EQUAL:
					return v.compareTo(t1) == 0;
				case NOT_EQUAL:
					return v.compareTo(t1) != 0;
				case GREATER_THAN:
					return v.compareTo(t1) > 0;
				case LESS_THAN:
					return v.compareTo(t1) < 0;
				case GREATER_OR_EQUAL:
					return v.compareTo(t1) >= 0;
				case LESS_OR_EQUAL:
					return v.compareTo(t1) <= 0;
				default:
					return true;
			}
		}
	}
	private class TextLengthValidator implements Validator {
		private int i1 = f1 == null ? 0 : Integer.parseInt(f1);
		private int i2 = f2 == null ? 0 : Integer.parseInt(f2);

		public boolean validate(String value) throws Exception {
			int v = value.length();
			switch (op) {
				case BETWEEN:
					return v >= i1 && v <= i2;
				case NOT_BETWEEN:
					return v <= i1 && v >= i2;
				case EQUAL:
					return v == i1;
				case NOT_EQUAL:
					return v == i1;
				case GREATER_THAN:
					return v > i1;
				case LESS_THAN:
					return v < i1;
				case GREATER_OR_EQUAL:
					return v >= i1;
				case LESS_OR_EQUAL:
					return v <= i1;
				default:
					return true;
			}
		}
	}
	private class FormulaValidator implements Validator {
		public boolean validate(String value) throws Exception {
			//Not implement
			return true;
		}
	}

	public String toJson() {
		return new Gson().toJson(this);
	}

	public static InputRule fromJson(String json) {
		InputRule ret = new Gson().fromJson(json, InputRule.class);
		CellRangeAddressList list = new CellRangeAddressList();
		for (String str : ret.regionsStr) {
			list.addCellRangeAddress(CellRangeAddress.valueOf(str));
		}
		ret.regions = list;
		return ret;
	}

	private static boolean compare(String s1, String s2) {
		if (s1 == null) return s2 == null;
		return s1.equals(s2);
	}

	public boolean ruleEquals(InputRule rule) {
		return 
			this.empty == rule.empty &&
			compare(this.errTitle, rule.errTitle) &&
			compare(this.errText, rule.errText) &&
			this.errStyle == rule.errStyle &&
			compare(this.pmTitle, rule.pmTitle) &&
			compare(this.pmText, rule.pmText) &&
			Arrays.equals(this.list, rule.list) &&
			compare(this.f1, rule.f1) &&
			compare(this.f2, rule.f2) &&
			this.op == rule.op &&
			this.vt == rule.vt;
	}

	@Override
	public String toString() { return toJson();}

	private static String[] buildList(Sheet sheet, String str) {
		if (str == null || str.length() == 0) {
			return new String[0];
		}
		if (str.length() > 2 && str.charAt(0) == '"' && str.charAt(str.length() - 1) == '"') {
			str = str.substring(1, str.length() - 1);
		}
		if (str.indexOf(',') != -1) {
			String[] ret = str.split(",");
			for (int i=0; i<ret.length; i++) {
				ret[i] = ret[i].trim();
			}
			return ret;
		}
		String sheetName = null;
		int sheetIndex = str.indexOf("!");
		if (sheetIndex != -1) {
			sheetName = str.substring(0, sheetIndex);
			str = str.substring(sheetIndex + 1);
		}
		if (sheetName != null && !sheetName.equals(sheet.getSheetName())) {
			sheet = sheet.getWorkbook().getSheet(sheetName);
		}
		List<String> list = new ArrayList<String>();
		try {
			CellValueHelper helper = new CellValueHelper(sheet.getWorkbook(), true);
			AreaReference area = new AreaReference(str);
			for (CellReference cRef : area.getAllReferencedCells()) {
				Cell cell = ExcelUtils.getCell(sheet, cRef.getRow(), cRef.getCol());
				String value = cell == null ? null : helper.getFormattedValue(cell).getValue();
				if (value != null && value.length() > 0) {
					list.add(value);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			list.add(str);
		}
		String[] ret = new String[list.size()];
		return (String[])list.toArray(ret);
	}
}