package jp.co.flect.excel2canvas;

import java.awt.Color;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

/**
 * Wrapper class of Excel color.
 */
public class ExcelColor {
	
	private static final short NO_COLOR = IndexedColors.AUTOMATIC.getIndex();
	
	private Workbook workbook;
	
	public ExcelColor(Workbook workbook) {
		this.workbook = workbook;
	}
	
	public Color getFillBackgroundColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		if (style instanceof XSSFCellStyle) {
			return getColor(style.getFillBackgroundColorColor());
		} else {
			return getColor(style.getFillBackgroundColor());
		}
	}
	
	public Color getFillForegroundColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		if (style instanceof XSSFCellStyle) {
			return getColor(style.getFillForegroundColorColor());
		} else {
			return getColor(style.getFillForegroundColor());
		}
	}
	
	public Color getLeftBorderColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		if (style instanceof XSSFCellStyle) {
			XSSFCellStyle xcs = (XSSFCellStyle)style;
			return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.LEFT));
		} else {
			return getColor(style.getLeftBorderColor());
		}
	}
	
	public Color getRightBorderColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		if (style instanceof XSSFCellStyle) {
			XSSFCellStyle xcs = (XSSFCellStyle)style;
			return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.RIGHT));
		} else {
			return getColor(style.getRightBorderColor());
		}
	}
	
	public Color getTopBorderColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		if (style instanceof XSSFCellStyle) {
			XSSFCellStyle xcs = (XSSFCellStyle)style;
			return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.TOP));
		} else {
			return getColor(style.getTopBorderColor());
		}
	}
	
	public Color getBottomBorderColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		if (style instanceof XSSFCellStyle) {
			XSSFCellStyle xcs = (XSSFCellStyle)style;
			return getColor(xcs.getBorderColor(XSSFCellBorder.BorderSide.BOTTOM));
		} else {
			return getColor(style.getBottomBorderColor());
		}
	}
	
	public Color getFontColor(Cell cell) {
		CellStyle style = cell.getCellStyle();
		if (style == null) {
			return null;
		}
		return getFontColor(this.workbook.getFontAt(style.getFontIndex()));
	}
	
	public Color getFontColor(Font font) {
		return font instanceof XSSFFont ? 
			getColor(((XSSFFont)font).getXSSFColor()) :
			getColor(font.getColor());
	}
	
	private static Color getColor(org.apache.poi.ss.usermodel.Color c) {
		if (c == null) {
			return null;
		} else if (c instanceof HSSFColor) {
			HSSFColor hc = (HSSFColor)c;
			short[] rgb = hc.getTriplet();
			return new Color(rgb[0], rgb[1], rgb[2]);
		} else if (c instanceof XSSFColor) {
			XSSFColor xc = (XSSFColor)c;
			byte[] data = null;
			if (xc.getTint() != 0.0) {
				data = getRgbWithTint(xc);
				byte[] argb = xc.getARgb();
			} else {
				data = xc.getARgb();
			}
			if (data == null) {
				return null;
			}
			int idx = 0;
			int alpha = 255;
			if (data.length == 4) {
				alpha = data[idx++] & 0xFF;
			}
			int r = data[idx++] & 0xFF;
			int g = data[idx++] & 0xFF;
			int b = data[idx++] & 0xFF;
			return new Color(r, g, b, alpha);
		} else {
			throw new IllegalStateException();
		}
	}
	
	public Color getColor(short colorIndex) {
		if (colorIndex == NO_COLOR) {
			return null;
		}
		CellStyle style = this.workbook.getCellStyleAt((short)0);
		short temp = style.getFillForegroundColor();
		try {
			style.setFillForegroundColor(colorIndex);
			org.apache.poi.ss.usermodel.Color c = style.getFillForegroundColorColor();
			return getColor(c);
		} finally {
			style.setFillForegroundColor(temp);
		}
	}
	
	private static byte[] getRgbWithTint(XSSFColor c) {
		byte[] rgb = c.getCTColor().getRgb();
		double tint = c.getTint();
		if (rgb != null && tint != 0.0) {
			if(rgb.length == 4) {
				byte[] tmp = new byte[3];
				System.arraycopy(rgb, 1, tmp, 0, 3);
				rgb = tmp;
			}
			for (int i=0; i<rgb.length; i++) {
				int lum = rgb[i] & 0xFF;
				double d = sRGB_to_scRGB(lum / 255.0);
				d = tint > 0 ? d * (1.0 - tint) + tint : d * (1 + tint);
				d = scRGB_to_sRGB(d);
				rgb[i] = (byte)Math.round(d * 255.0);
			}
		}
		return rgb;
	}
	
	private static double sRGB_to_scRGB(double value) {
		if (value < 0.0) {
			return 0.0;
		}
		if (value <= 0.04045) {
			return value /12.92;
		}
		if (value <= 1.0) {
			return Math.pow(((value + 0.055) / 1.055), 2.4);
		}
		return 1.0;
	}
	
	private static double scRGB_to_sRGB(double value) {
		if (value < 0.0) {
			return 0.0;
		}
		if (value <= 0.0031308) {
			return value * 12.92;
		}
		if (value < 1.0) {
			return 1.055 * (Math.pow(value, (1.0 / 2.4))) - 0.055;
		}
		return 1.0;
	}
	
}
