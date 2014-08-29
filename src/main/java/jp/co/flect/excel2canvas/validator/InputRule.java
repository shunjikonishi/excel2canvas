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
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.hssf.record.cf.CellRangeUtil;
import org.joda.time.DateTime;
import org.joda.time.LocalTime;
import org.joda.time.format.DateTimeFormatterBuilder;
import org.joda.time.format.DateTimeFormatter;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeParser;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.List;
import java.util.ArrayList;
import java.text.ParseException;
import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import org.w3c.dom.Node;
import org.w3c.dom.Element;

import jp.co.flect.excel2canvas.ExcelUtils;
import jp.co.flect.excel2canvas.CellValueHelper;

public class InputRule {

	private static final DateTimeFormatter DATE_FORMATTER;

	static {
		DateTimeParser[] parsers = {
			DateTimeFormat.forPattern("yyyy-MM-dd").getParser(),
			DateTimeFormat.forPattern("yyyy/MM/dd").getParser()
		};
		DATE_FORMATTER = new DateTimeFormatterBuilder().append(null, parsers).toFormatter();
	}

	public static DateTime parseDateTime(String str) {
		try {
			return DATE_FORMATTER.parseDateTime(str);
		} catch (IllegalArgumentException e) {
			try {
				double n = Double.parseDouble(str);
				return new DateTime(DateUtil.getJavaDate(n));
			} catch (NumberFormatException e2) {
				//Ignore
			}
			throw e;//Rethrow IllegalArgumentException
		}
	}
	//type="list" only
	public static InputRule fromDataValidationNode(Sheet sheet, Element el) {
		String localName = el.getLocalName();
		String type = el.getAttribute("type");
		if (!"dataValidation".equals(localName) || !"list".equals(type)) {
			return null;
		}
		return new InputRule(sheet, el);
	}

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
		if (dv.getShowErrorBox()) {
			errTitle = dv.getErrorBoxTitle();
			errText = dv.getErrorBoxText();
			errStyle = dv.getErrorStyle();
		}
		if (dv.getShowPromptBox()) {
			pmTitle = dv.getPromptBoxTitle();
			pmText = dv.getPromptBoxText();
		}

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

	/*
	<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="rrr" error="xxxx" promptTitle="sss" prompt="nnn">
	  <x14:formula1>
	    <xm:f>Sheet2!$A$1:$A$3</xm:f>
	  </x14:formula1>
	  <xm:sqref>B5</xm:sqref>
	</x14:dataValidation>
	x14:dataValidations>
	*/
	private InputRule(Sheet sheet, Element el) {
		empty = "1".equals(el.getAttribute("allowBlank"));
		if ("1".equals(el.getAttribute("showErrorMessage"))) {
			errTitle = checkNull(el.getAttribute("errorTitle"));
			errText = checkNull(el.getAttribute("error"));
			//errStyle = ???
		}
		if ("1".equals(el.getAttribute("showInputMessage"))) {
			pmTitle = checkNull(el.getAttribute("promptTitle"));
			pmText = checkNull(el.getAttribute("prompt"));
		}
		String region = null;
		Node node = el.getFirstChild();
		while (node != null) {
			String name = node.getLocalName();
			if ("formula1".equals(name)) {
				f1 = getChildText((Element)node);
			} else if ("formula2".equals(name)) {
				f2 = getChildText((Element)node);
			} else if ("sqref".equals(name)) {
				region = getChildText((Element)node);
			}
			node = node.getNextSibling();
		}
		if (f1 == null || region == null) {
			throw new IllegalArgumentException("Invalid element: " + el.getNodeName());
		}
		regionsStr = new String[1];
		regionsStr[0] = region;
		regions = new CellRangeAddressList();
		regions.addCellRangeAddress(CellRangeAddress.valueOf(region));

		//op = ???
		vt = DataValidationConstraint.ValidationType.LIST;
		list = buildList(sheet, f1);
	}

	public boolean getAllowEmpty() { return empty;}
	public String getErrorTitle() { return errTitle;}
	public String getErrorText() { return errText;}
	public int getErrorStyle() { return errStyle;}
	public boolean hasOwnErrorText() { return errTitle != null || errText != null;}

	public String getPromptTitle() { return pmTitle;}
	public String getPromptText() { return pmText;}
	public boolean hasPrompt() { return pmTitle != null || pmText != null;}

	public String[] getRegions() { return regionsStr;}
	public String[] getList() { return list;}

	public String getFormula1() { return f1;}
	public String getFormula2() { return f2;}

	public int getOperator() { return op;}
	public int getValidationType() { return vt;}

	public String getValidationTerm() {
		if (vt == ANY || vt == FORMULA) {
			return null;
		}
		if (vt == LIST) {
			return Arrays.toString(list);
		}
		String vStr = vt == TEXT_LENGTH ? "length" : "value";
		if (f1 == null) {
			return null;
		}
		switch (op) {
			case BETWEEN:
				return f1 + " <= " + vStr + " <= " + f2;
			case NOT_BETWEEN:
				return vStr + " < " + f1 + " or " + f2 + " < " + vStr;
			case EQUAL:
				return vStr + " == " + f1;
			case NOT_EQUAL:
				return vStr + " != " + f1;
			case GREATER_THAN:
				return vStr + " > " + f1;
			case LESS_THAN:
				return vStr + " < " + f1;
			case GREATER_OR_EQUAL:
				return vStr + " >= " + f1;
			case LESS_OR_EQUAL:
				return vStr + " <= " + f1;
			default:
				return null;
		}
	}

	public boolean isOverlapped(Name name) {
		String ref = name.getRefersToFormula();
		int idx = ref.indexOf('!');
		if (idx != -1) {
			ref = ref.substring(idx + 1);
		}
		AreaReference area = new AreaReference(ref);
		CellReference topLeft = area.getFirstCell();
		CellReference bottomRight = area.getLastCell();
		CellRangeAddress cra = new CellRangeAddress(
			topLeft.getRow(), bottomRight.getRow(), 
			topLeft.getCol(), bottomRight.getCol()
		);
		return isOverlapped(cra);
	}

	public boolean isOverlapped(CellRangeAddress cra) {
		for (CellRangeAddress c : this.regions.getCellRangeAddresses()) {
			if (CellRangeUtil.intersect(cra, c) != CellRangeUtil.NO_INTERSECTION) {
				return true;
			}
		}
		return false;
	}

	public void validate(String value) throws Exception {
		if (value == null || value.length() == 0) {
			if (empty) {
				return;
			}
			throw new EmptyValueException("Value is required.");
		}
		if (f1 == null) {
			return;
		}
		if (validator == null) {
			validator = createValidator();
		}
		try {
			if (!validator.validate(value)) {
				throw new InvalidValueException("Invalid value.");
			}
		} catch (NumberFormatException e) {
			throw new InvalidValueException(e);
		} catch (ParseException e) {
			throw new InvalidValueException(e);
		} catch (IllegalArgumentException e) {
			throw new InvalidValueException(e);
		}
	}

	private Validator createValidator() {
		switch (vt) {			case ANY:
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
					return v.compareTo(f1) < 0 || v.compareTo(f2) > 0;
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
					return v < l1 || v > l2;
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
					return v.compareTo(d1) < 0 || v.compareTo(d2) > 0;
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
		private DateTime d1 = f1 == null ? null : parseDateTime(f1);
		private DateTime d2 = f2 == null ? null : parseDateTime(f2);

		public boolean validate(String value) throws Exception {
			DateTime v = parseDateTime(value);
			switch (op) {
				case BETWEEN:
					return v.compareTo(d1) >= 0 && v.compareTo(d2) <= 0;
				case NOT_BETWEEN:
					return v.compareTo(d1) < 0 || v.compareTo(d2) > 0;
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
					return v.compareTo(t1) < 0 || v.compareTo(t2) > 0;
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
					return v < i1 || v > i2;
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
		setupRegions(ret);
		return ret;
	}

	private static void setupRegions(InputRule rule) {
		CellRangeAddressList list = new CellRangeAddressList();
		for (String str : rule.regionsStr) {
			list.addCellRangeAddress(CellRangeAddress.valueOf(str));
		}
		rule.regions = list;
	}

	public static List<InputRule> fromJsonArray(String json) {
		List<InputRule> ret = new Gson().fromJson(json, new TypeToken<List<InputRule>>() {}.getType());
		for (InputRule rule : ret) {
			setupRegions(rule);
		}
		return ret;
	}

	private static boolean compare(String s1, String s2) {
		if (s1 == null) return s2 == null;
		return s1.equals(s2);
	}

	private static String checkNull(String s) {
		return s != null && s.length() > 0 ? s : null;
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
		//Comma separated literal
		if (str.indexOf(',') != -1) {
			String[] ret = str.split(",");
			for (int i=0; i<ret.length; i++) {
				ret[i] = ret[i].trim();
			}
			return ret;
		}

		List<String> list = new ArrayList<String>();
		//Name reference
		Name name = sheet.getWorkbook().getName(str);
		if (name != null) {
			str = name.getRefersToFormula();
		}
		//Cell reference
		String refStr = str;
		int sheetIndex = str.indexOf("!");
		if (sheetIndex != -1) {
			String sheetName = str.substring(0, sheetIndex);
			if (!sheetName.equals(sheet.getSheetName())) {
				sheet = sheet.getWorkbook().getSheet(sheetName);
			}
			refStr = str.substring(sheetIndex + 1);
		}
		try {
			AreaReference area = new AreaReference(refStr);
			CellValueHelper helper = new CellValueHelper(sheet.getWorkbook(), true);
			for (CellReference cRef : area.getAllReferencedCells()) {
				Cell cell = ExcelUtils.getCell(sheet, cRef.getRow(), cRef.getCol());
				String value = cell == null ? null : helper.getFormattedValue(cell).getValue();
				if (value != null && value.length() > 0) {
					list.add(value);
				}
			}
		} catch (Exception e) {
			list.add(str);
		}
		String[] ret = new String[list.size()];
		return (String[])list.toArray(ret);
	}

	private static String getChildText(Element el) {
		StringBuilder buf = new StringBuilder();
		buildChildText(el, buf);
		return buf.toString();
	}

	private static void buildChildText(Element el, StringBuilder buf) {
		Node node = el.getFirstChild();
		while (node != null) {
			int type = node.getNodeType();
			switch (type) {
				case Node.TEXT_NODE:
				case Node.CDATA_SECTION_NODE:
					String str = node.getNodeValue();
					if (!isWhitespace(str)) {
						buf.append(str);
					}
					break;
				case Node.ELEMENT_NODE:
					buildChildText((Element)node, buf);
					break;
			}
			node = node.getNextSibling();
		}
	}

	private static boolean isWhitespace(String s) {
		if (s == null || s.length() == 0) {
			return true;
		}
		for (int i=0; i<s.length(); i++) {
			char c = s.charAt(i);
			if (c == ' ' || c == '\t' || c == '\r' || c == '\n') {
				continue;
			}
			return false;
		}
		return true;
	}

	public static class InvalidValueException extends Exception {
		public InvalidValueException(String msg) {
			super(msg);
		}

		public InvalidValueException(Throwable e) {
			super(e);
		}
	}

	public static class EmptyValueException extends Exception {
		public EmptyValueException(String msg) {
			super(msg);
		}
	}

	//For Test
	static InputRule forTest(int vt, int op, String f1, String f2) {
		return new InputRule(vt, op, f1, f2);
	}
	private InputRule(int vt, int op, String f1, String f2) {
		this.vt = vt;
		this.op = op;
		this.f1 = f1;
		this.f2 = f2;
	}

}