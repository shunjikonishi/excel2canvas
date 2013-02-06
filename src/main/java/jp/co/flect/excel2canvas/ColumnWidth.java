package jp.co.flect.excel2canvas;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.helpers.ColumnHelper;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCol;

/**
 * Calculate cell width of worksheet.
 */
public class ColumnWidth {
	
	private static final double SCALE_ARIAL    = 36.5625;
	private static final double SCALE_MSGOTHIC = 32.0;
	
	private static final int DEFAULT_ARIAL     = 64;
	private static final int DEFAULT_MSGOTHIC  = 72;
	
	private static final int DEFAULT_WIDTH    = 2048;
	
	private static double getScale(Font font) {
		String name = font.getFontName();
		int pt = font.getFontHeightInPoints();
		if ("ＭＳ Ｐゴシック".equals(name) && pt == 11) return SCALE_MSGOTHIC;
		if ("Arial".equals(name) && pt == 10) return SCALE_ARIAL;
		
		System.out.println("Unknown font: " + name + ", " + pt);
		return SCALE_MSGOTHIC;
	}
	
	private double scale;
	
	public ColumnWidth(Workbook workbook) {
		this.scale = getScale(workbook.getFontAt((short)0));
	}
	
	public int getColumnWidth(Sheet sheet, int col) {
		int colWidth = sheet.getColumnWidth(col);
		if (colWidth == DEFAULT_WIDTH) {
			boolean bDefault = true;
			if (sheet instanceof XSSFSheet) {
				CTCol ctCol = ((XSSFSheet)sheet).getColumnHelper().getColumn(col, false);
				if (ctCol != null && ctCol.isSetWidth()) {
					bDefault = false;
				}
			}
			if (bDefault) {
				if (scale == SCALE_ARIAL) {
					return DEFAULT_ARIAL;
				} else {
					return DEFAULT_MSGOTHIC;
				}
			}
		}
		double d = colWidth / scale;
		return (int)Math.round(d);
	}
	
}
