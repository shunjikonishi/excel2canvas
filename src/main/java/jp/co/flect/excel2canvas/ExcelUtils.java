package jp.co.flect.excel2canvas;

import java.awt.Point;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.BitSet;
import java.util.HashSet;
import java.util.Set;
import java.util.HashMap;
import java.util.Map;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import org.w3c.dom.Element;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;

import jp.co.flect.excel2canvas.validator.InputRule;
import jp.co.flect.excel2canvas.validator.ExtDataValidationExtractor;

/**
 * Utility
 */
public class ExcelUtils {
	
	/** POIのWorkbookFactoryはInputStreamをクローズしてくれない */
	public static Workbook createWorkbook(File file) throws IOException, InvalidFormatException {
		FileInputStream is = new FileInputStream(file);
		try {
			return WorkbookFactory.create(is);
		} finally {
			is.close();
		}
	}
	
	/** POIのWorkbookFactoryはInputStreamをクローズしてくれない */
	public static Workbook createWorkbook(InputStream is) throws IOException, InvalidFormatException {
		try {
			return WorkbookFactory.create(is);
		} finally {
			is.close();
		}
	}
	
	public static Point nameToPoint(String name) {
		int i=0;
		for (i=0; i<name.length(); i++) {
			char c = name.charAt(i);
			if (Character.isDigit(c)) {
				break;
			}
		}
		return new Point(nameToColumn(name.substring(0, i)), Integer.parseInt(name.substring(i)) - 1);
	}
	
	public static int nameToColumn(String name) {
		int column = -1;
		for (int i = 0; i < name.length(); ++i) {
			int c = name.charAt(i);
			column = (column + 1) * 26 + c - 'A';
		}
		return column;
	}
	
	public static String pointToName(Point p) {
		return pointToName(p.x, p.y);
	}
	
	public static String pointToName(int x, int y) {
		StringBuilder buf = new StringBuilder();
		if (x > 25) {
			int a = x / 26 - 1;
			int b = x % 26;
			buf.append((char)('A' + a)).append((char)('A' + b));
		} else {
			buf.append((char)('A' + x));
		}
		buf.append(y + 1);
		return buf.toString();
	}
	
	public static int getRowHeight(Sheet sheet, int row) {
		Row r = sheet.getRow(row);
		int h = r == null ? sheet.getDefaultRowHeight() : r.getHeight();
		return h / 15;
	}
	
	/**
	 * POIのDateUtilにあるメソッドが日本語を含む日付書式を正しく扱ってくれないので自力実装
	 */
	public static boolean isCellDateFormatted(Cell cell) {
		if (cell == null) {
			return false;
		}
		return isCellDateFormatted(cell, cell.getNumericCellValue());
	}
	
	/**
	 * POIのDateUtilにあるメソッドが日本語を含む日付書式を正しく扱ってくれないので自力実装
	 */
	public static boolean isCellDateFormatted(Cell cell, double d) {
		if (cell == null) {
			return false;
		}
		boolean bDate = false;
		if (DateUtil.isValidExcelDate(d)) {
			CellStyle style = cell.getCellStyle();
			if (style == null) {
				return false;
			}
			int i = style.getDataFormat();
			String f = style.getDataFormatString();
			bDate = isADateFormat(i, f);
		}
		return bDate;
	}

	public static boolean isNumericStyle(CellStyle style) {
		return style != null && isNumericStyle(style.getDataFormatString());
	}

	public static boolean isNumericStyle(String str) {
		if (str == null || str.length() == 0) {
			return false;
		}
		boolean skip = false;
		for (int i=0; i<str.length(); i++) {
			char c = str.charAt(i);
			if (skip) {
				if (c == ']' || c == '"') {
					skip = false;
				}
				continue;
			}
			switch (c) {
				case '@':
				case '0':
				case '#':
				case '\\':
				case '(':
				case ')':
				case ',':
				case ';':
				case '_':
				case '-':
				case ' ':
					break;
				case '[':
				case '"':
					skip = true;
					break;
				default:
					return false;
			}
		}
		return true;
	}
	
	/**
	 * フォーマットが日付書式であるかどうかを判定します
	 * @See http://support.microsoft.com/kb/883199/ja
	 */
	public static boolean isADateFormat(int formatIndex, String formatString) {
		//builtin - org.apache.poi.ss.usermodel.BuiltinFormats
		if (DateUtil.isInternalDateFormat(formatIndex)) {
			return true;
		}
		if (formatString == null) {
			return false;
		}
		int len = formatString.length();
		for (int i=0; i<len; i++) {
			char c = formatString.charAt(i);
			switch (c) {
				case '0':
				case '#':
					return false;//数値用の書式文字列
				case 'y':
				case 'e':
				case 'g':
				case 'm':
				case 'd':
				case 'a':
				case 'h':
				case 's':
					return true;//日付用の書式文字列
				case 'r':
					if (i == 0 && formatString.startsWith("reserved")) {
						return false;
					}
					break;
				case 'G'://General
				{
					int idx = formatString.indexOf("General", i);
					if (idx == i) {
						return false;
					}
					break;
				}
				case '\\'://次の文字をエスケープ
					i++;
					break;
				case '"'://文字列
				{
					int idx = formatString.indexOf(']', i+1);
					if (idx == -1) {
						idx = len;
					}
					i = idx;
					break;
				}
				case '[':
				{
					int idx = formatString.indexOf(']', i+1);
					if (idx == -1) {
						idx = len;
					} else if (idx == i+2) {
						char c2 = formatString.charAt(i+1);
						if (c2 == 'h' || c2 == 'm' || c2 == 's') {
							return true;
						}
					} else if (idx == i+3) {
						char c2 = formatString.charAt(i+1);
						char c3 = formatString.charAt(i+2);
						if (c2 == c3 && c2 == 'h' || c2 == 'm' || c2 == 's') {
							return true;
						}
					}
					i = idx;
					break;
				}
			}
		}
		return false;
	}
	/*
	// 2012/06/25以前に使用していたロジック
	// 日付の書式文字が2文字以上現れたら日付書式と判定する(本当にそれで良いのかは？？？)
	public static boolean isADateFormat(int formatIndex, String formatString) {
		if (DateUtil.isADateFormat(formatIndex, formatString)) {
			return true;
		}
		//builtin - org.apache.poi.ss.usermodel.BuiltinFormats
		if (formatIndex <= 0x31) {
			return false;
		}
		if (formatString == null) {
			return false;
		}
		
		int flag = -1;
		int flagCnt = 0;
		final String dateStr = "geymdhsa";
		for (int i=0; i<formatString.length(); i++) {
			char c = formatString.charAt(i);
			if (c == '0' || c == '#') {
				//数値書式と思われる
				return false;
			} else if (c == '[') {
				//[]内はスキップ
				int n = formatString.indexOf(']', i+1);
				if (n != -1) {
					i = n;
				}
				continue;
			}
			int idx = dateStr.indexOf(c);
			if (idx != -1) {
				if (flag != -1 && flag != idx) {
					return true;
				}
				flag = idx;
				flagCnt++;
			}
		}
		return flagCnt == formatString.length();
	}
	*/
	
	/**
	 * 書式文字列を「;」で分割します。
	 */
	public static String[] splitFormat(String str) {
		if (str.indexOf(';') == -1) {
			String[] ret = new String[1];
			ret[0] = str;
			return ret;
		}
		List<String> list = new ArrayList<String>();
		int spos = 0;
		int len = str.length();
		for (int i=0; i<len; i++) {
			char c = str.charAt(i);
			switch (c) {
				case ';':
					list.add(str.substring(spos, i));
					spos = i+1;
					break;
				case '\\':
				case '_':
				case '*':
					i++;
					break;
				case '"':
				{
					int idx = str.indexOf('"', i+1);
					if (idx != -1) {
						i = idx;
					}
					break;
				}
				case '[':
				{
					int idx = str.indexOf(']', i+1);
					if (idx != -1) {
						i = idx;
					}
					break;
				}
			}
		}
		if (spos < len - 1) {
			list.add(str.substring(spos));
		}
		String[] ret = new String[list.size()];
		return (String[])list.toArray(ret);
	}
	
	public static Cell getOrCreateCell(Sheet sheet, int rc, int cc) {
		Row row = sheet.getRow(rc);
		if (row == null) {
			row = sheet.createRow(rc);
		}
		
		Cell cell = row.getCell(cc);
		if (cell == null) {
			cell = row.createCell(cc);
		}
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			style = sheet.getColumnStyle(cc);
			if (style == null) {
				style = row.getRowStyle();
			}
			if (style != null) {
				cell.setCellStyle(style);
			}
		}
		return cell;
	}
	
	public static Cell getCell(Sheet sheet, int rc, int cc) {
		Row row = sheet.getRow(rc);
		if (row == null) {
			return null;
		}
		return row.getCell(cc);
	}

	/**
	 * CellStyleに何も情報が設定されていない場合にtrueを返す
	 */
	public static boolean isNoneStyle(CellStyle style) {
		if (style == null) return true;
		
		if (style.getDataFormat() != 0) return false;
		if (style.getFillPattern() != CellStyle.NO_FILL) return false;
		if (style.getFillBackgroundColor() != IndexedColors.AUTOMATIC.getIndex()) return false;
		if (style.getFillForegroundColor() != IndexedColors.AUTOMATIC.getIndex()) return false;
		if (style.getFontIndex() != 0) return false;
		if (style.getIndention() != 0) return false;
		if (style.getRotation() != 0) return false;
		if (style.getAlignment() != CellStyle.ALIGN_GENERAL) return false;
		if (style.getVerticalAlignment() != CellStyle.VERTICAL_CENTER) return false;
		if (style.getBorderTop() != CellStyle.BORDER_NONE) return false;
		if (style.getBorderLeft() != CellStyle.BORDER_NONE) return false;
		if (style.getBorderRight() != CellStyle.BORDER_NONE) return false;
		if (style.getBorderBottom() != CellStyle.BORDER_NONE) return false;
		if (style.getWrapText()) return false;
		if (style.getHidden()) return false;
		if (!style.getLocked()) return false;
		
		return true;
	}
	
	public static void resultSetToExcel(ResultSet rs, File f) throws IOException, SQLException {
		SXSSFWorkbook wb = new SXSSFWorkbook(100);
		try {
			Sheet sheet = wb.createSheet();
			
			ResultSetMetaData meta = rs.getMetaData();
			int rowNum = 0;
			int colLen = meta.getColumnCount();
			boolean[] types = new boolean[colLen];
			Row row = sheet.createRow(rowNum++);
			for (int i=0; i<colLen; i++) {
				Cell cell = row.createCell(i);
				cell.setCellValue(meta.getColumnLabel(i+1));
				int n = meta.getColumnType(i+1);
				types[i] = n == Types.BIGINT ||
				           n == Types.DECIMAL ||
				           n == Types.DOUBLE ||
				           n == Types.FLOAT ||
				           n == Types.INTEGER ||
				           n == Types.NUMERIC ||
				           n == Types.REAL ||
				           n == Types.SMALLINT ||
				           n == Types.TINYINT;
			}
			
			while (rs.next()) {
				row = sheet.createRow(rowNum++);
				for (int i=0; i<colLen; i++) {
					Cell cell = row.createCell(i);
					if (types[i]) {
						cell.setCellValue(rs.getDouble(i+1));
					} else {
						cell.setCellValue(rs.getString(i+1));
					}
				}
			}
			FileOutputStream os = new FileOutputStream(f);
			try {
				wb.write(os);
			} finally {
				os.close();
			}
		} finally {
			wb.dispose();
		}
	}

	public static CellStyle createDefaultLabelStyle(Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		font.setFontHeightInPoints((short)12);
		style.setFont(font);
		style.setBorderBottom(CellStyle.BORDER_DOUBLE);
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		return style;
	}
	
	public static CellStyle createDateStyle(Workbook workbook, String format) {
		CellStyle style = workbook.createCellStyle();
		DataFormat df = workbook.createDataFormat();
		style.setDataFormat(df.getFormat(format));
		return style;
	}
	
	public static List<NamedCellInfo> createNamedCellList(Workbook workbook) {
		return createNamedCellList(workbook, null);
	}
	
	public static List<NamedCellInfo> createNamedCellList(Workbook workbook, String targetSheet) {
		Set<Sheet> sheetSet = new HashSet<Sheet>();
		List<TempName> tempList = new ArrayList<TempName>();
		for (int i=0; i<workbook.getNumberOfNames(); i++) {
			//isDeleted, isFunctionNameは正しい値を返さないことがあり、
			//その場合getSheetNameの実行またはAreaReferenceの生成でエラーとなる
			//それらは正しい名前参照ではないのでNamedCellInfoのListには含めない
			Name name = workbook.getNameAt(i);
			try {
				if (name.getSheetName() == null) {
					continue;
				}
				if (name.isFunctionName() || name.isDeleted()) {
					continue;
				}
				if (targetSheet != null && !targetSheet.equals(name.getSheetName())) {
					continue;
				}
				Sheet sheet = workbook.getSheet(name.getSheetName());
				if (sheet == null) {
					continue;
				}
				sheetSet.add(sheet);
				
				AreaReference area = new AreaReference(name.getRefersToFormula());
				tempList.add(new TempName(name, area, sheet.getNumMergedRegions()));
			} catch (Exception e) {
				//Ignore
				System.out.println("Ignore error of ExcelUtils#createNamedList: " + name.getNameName() + ", " + e.toString());
			}
		}
		for (Sheet sheet : sheetSet) {
			for (int i=0; i<sheet.getNumMergedRegions(); i++) {
				CellRangeAddress mr = sheet.getMergedRegion(i);
				for (TempName temp : tempList) {
					if (temp.getSheetName().equals(sheet.getSheetName()) && temp.isLap(mr)) {
						temp.relatedMergedRegion.set(i);
					}
				}
			}
		}
		List<NamedCellInfo> ret = new ArrayList<NamedCellInfo>(tempList.size());
		for (TempName temp : tempList) {
			Sheet sheet = workbook.getSheet(temp.getSheetName());
			List<CellReference> cellList = new ArrayList<CellReference>();
			BitSet bs = temp.relatedMergedRegion;
			
			CellReference[] cells = temp.area.getAllReferencedCells();
			for (CellReference cell : cells) {
				boolean add = true;
				for (int i=bs.nextSetBit(0); i>=0; i=bs.nextSetBit(i+1)) {
					CellRangeAddress mr = sheet.getMergedRegion(i);
					if (mr.isInRange(cell.getRow(), cell.getCol())) {
						if (mr.getFirstRow() != cell.getRow() || mr.getFirstColumn() != cell.getCol()) {
							add = false;
						}
						break;
					}
				}
				if (add) {
					cellList.add(cell);
				}
			}
			if (cellList.size() > 0) {
				ret.add(new NamedCellInfo(temp.name, cellList));
			}
		}
		return ret;
	}

	public static Map<String, List<InputRule>> getInputRules(Workbook workbook) {
		Map<String, List<InputRule>> ret = new HashMap<String, List<InputRule>>();
		for (int i=0; i<workbook.getNumberOfSheets(); i++) {
			Sheet sheet = workbook.getSheetAt(i);
			ret.put(sheet.getSheetName(), getInputRules(sheet));
		}
		return ret;
	}

	public static List<InputRule> getInputRules(Sheet sheet) {
		List<InputRule> ret = new ArrayList<InputRule>();
		if (sheet instanceof XSSFSheet) {
			XSSFSheet xs = (XSSFSheet)sheet;
			for (XSSFDataValidation dv : xs.getDataValidations()) {
				ret.add(new InputRule(sheet, dv));
			}
			List<Element> extList = new ExtDataValidationExtractor().getDataValidationNode(xs);
			for (Element el : extList) {
				InputRule rule = InputRule.fromDataValidationNode(xs, el);
				if (rule != null) {
					ret.add(rule);
				}
			}
		}
		return ret;
	}
	
	private static class TempName {
		
		public Name name;
		public AreaReference area;
		public BitSet relatedMergedRegion;
		
		public TempName(Name name, AreaReference area, int mrCount) {
			this.name = name;
			this.area = area;
			this.relatedMergedRegion = new BitSet(mrCount);
		}
		
		public String getSheetName() {
			return this.area.getFirstCell().getSheetName();
		}
		
		public boolean isLap(CellRangeAddress mr) {
			int minRow1 = mr.getFirstRow();
			int maxRow1 = mr.getLastRow();
			int minCol1 = mr.getFirstColumn();
			int maxCol1 = mr.getLastColumn();
			
			if (area.isSingleCell()) {
				int row = area.getFirstCell().getRow();
				int col = area.getFirstCell().getCol();
				return minRow1 <= row && row <= maxRow1 && minCol1 <= col && col <= minCol1;
			}
			int minRow2 = Math.min(area.getFirstCell().getRow(), area.getLastCell().getRow());
			int maxRow2 = Math.max(area.getFirstCell().getRow(), area.getLastCell().getRow());
			int minCol2 = Math.min(area.getFirstCell().getCol(), area.getLastCell().getCol());
			int maxCol2 = Math.max(area.getFirstCell().getCol(), area.getLastCell().getCol());
			
			Rectangle r1 = new Rectangle(minCol1, minRow1, maxCol1 - minCol1 + 1, maxRow1 - minRow1 + 1);
			Rectangle r2 = new Rectangle(minCol2, minRow2, maxCol2 - minCol2 + 1, maxRow2 - minRow2 + 1);
			return r1.intersects(r2);
		}
	}
	
}
