package jp.co.flect.excel2canvas;

import java.awt.Color;
import java.awt.Point;
import java.awt.Rectangle;
import java.io.File;
import java.io.InputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Locale;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.Set;
import java.util.HashSet;
import java.util.Objects;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.hssf.usermodel.HSSFAnchor;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTDrawing;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTTwoCellAnchor;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTPicture;

import jp.co.flect.excel2canvas.chart.Chart;
import jp.co.flect.excel2canvas.chart.ChartFactory;
import jp.co.flect.excel2canvas.chart.Flotr2ChartFactory;

/**
 * Builder class of ExcelToCanvas.
 */
public class ExcelToCanvasBuilder {
	
	static {
		jp.co.flect.excel2canvas.functions.FunctionManager.registerAll();
	}
	
	private Locale locale;
	private FontManager fontManager = new FontManager();
	
	private Workbook workbook;
	private Sheet sheet;
	
	private ColumnWidth columnWidth;
	private ExcelColor exColor;
	
	private int maxCol;
	private int maxRow;
	private int writeCol;
	private int writeRow;
	
	private CellWrapper[][] cells;
	private CellValueHelper helper;
	
	private int readMaxRow = 0;
	private int readMaxCol = 0;
	private Rectangle readRange = null;
	
	private boolean includeEmptyStr = false;
	private boolean includeRawData = false;
	private boolean includeComment = false;
	private boolean includeChart = false;
	private boolean includePicture = false;
	private boolean includeHiddenCell = false;
	private ExpandChecker expandChecker = null;
	private Set<String> includeCells = null;
	
	private ChartFactory chartFactory = null;
	
	public ExcelToCanvasBuilder() {
		this(Locale.getDefault());
	}
	
	public ExcelToCanvasBuilder(Locale l) {
		this.locale = l;
	}

	public ExcelToCanvasMap buildAll(File f) throws IOException, InvalidFormatException {
		return buildAll(ExcelUtils.createWorkbook(f));
	}
	
	public ExcelToCanvasMap buildAll(InputStream is) throws IOException, InvalidFormatException {
		return buildAll(ExcelUtils.createWorkbook(is));
	}
	
	public ExcelToCanvasMap buildAll(Workbook workbook) throws IOException, InvalidFormatException {
		ExcelToCanvasMap map = new ExcelToCanvasMap();
		for (int i=0; i<workbook.getNumberOfSheets(); i++) {
			String name = workbook.getSheetName(i);
			map.put(name, build(workbook, name));
		}
		return map;
	}
	
	public ExcelToCanvas build(File f) throws IOException, InvalidFormatException {
		return build(ExcelUtils.createWorkbook(f));
	}
	
	public ExcelToCanvas build(InputStream is) throws IOException, InvalidFormatException {
		return build(ExcelUtils.createWorkbook(is));
	}
	
	public ExcelToCanvas build(File f, String sheetName) throws IOException, InvalidFormatException {
		return build(ExcelUtils.createWorkbook(f), sheetName);
	}
	
	public ExcelToCanvas build(InputStream is, String sheetName) throws IOException, InvalidFormatException {
		return build(ExcelUtils.createWorkbook(is), sheetName);
	}
	
	public ExcelToCanvas build(Workbook workbook) {
		return build(workbook, null);
	}
	
	public ExcelToCanvas build(Workbook workbook, String sheetName) {
		this.workbook = workbook;
		this.sheet = null;
		if (sheetName != null && sheetName.length() > 0) {
			this.sheet = workbook.getSheet(sheetName);
		} else {
			for (int i=0; i<workbook.getNumberOfSheets(); i++) {
				if (!workbook.isSheetHidden(i) && !workbook.isSheetVeryHidden(i)) {
					sheet = workbook.getSheetAt(i);
					break;
				}
			}
		}
		if (sheet == null) {
			throw new IllegalArgumentException(sheetName);
		}
		this.cells = null;
		this.helper = new CellValueHelper(this.workbook, true, this.locale);
		this.columnWidth = new ColumnWidth(this.workbook);
		this.exColor = new ExcelColor(this.workbook);
		this.maxCol = 0;
		this.maxRow = 0;
		this.writeCol = 0;
		this.writeRow = 0;
		
		return build();
	}
	
	public boolean isIncludeEmptyStr() { return this.includeEmptyStr;}
	public void setIncludeEmptyStr(boolean b) { this.includeEmptyStr = b;}
	
	public void addIncludeCell(String cellName) {
		if (this.includeCells == null) {
			this.includeCells = new HashSet<String>();
		}
		this.includeCells.add(cellName);
	}
	public void clearIncludeCells() {
		this.includeCells = null;
	}
	
	public boolean isIncludeRawData() { return this.includeRawData;}
	public void setIncludeRawData(boolean b) { this.includeRawData = b;}
	
	public boolean isIncludeComment() { return this.includeComment;}
	public void setIncludeComment(boolean b) { this.includeComment = b;}
	
	public boolean isIncludeChart() { return this.includeChart;}
	public void setIncludeChart(boolean b) { this.includeChart = b;}
	
	public boolean isIncludePicture() { return this.includePicture;}
	public void setIncludePicture(boolean b) { this.includePicture = b;}
	
	public boolean isIncludeHiddenCell() { return this.includeHiddenCell;}
	public void setIncludeHiddenCell(boolean b) { this.includeHiddenCell = b;}
	
	public ExpandChecker getExpandChecker() { return this.expandChecker;}
	public void setExpandChecker(ExpandChecker c) { this.expandChecker = c;}
	
	public int getReadMaxRow() { return this.readMaxRow;}
	public void setReadMaxRow(int n) { this.readMaxRow = n;}
	
	public int getReadMaxCol() { return this.readMaxCol;}
	public void setReadMaxCol(int n) { this.readMaxCol = n;}
	
	public ChartFactory getChartFactory() { return this.chartFactory;}
	public void setChartFactory(ChartFactory f) { this.chartFactory = f;}
	
	public FontManager getFontManager() { return this.fontManager;}
	public void setFontManager(FontManager m) { this.fontManager = m;}
	
	public String getReadRange() { 
		if (this.readRange == null) {
			return null;
		}
		String topLeft = ExcelUtils.pointToName(this.readRange.x, this.readRange.y);
		String bottomRight = ExcelUtils.pointToName(this.readRange.width, this.readRange.height);
		return topLeft + ":" + bottomRight;
	}
	
	public void setReadRange(String s) {
		if (s == null || s.length() == 0) {
			this.readRange = null;
		} else {
			String[] strs = s.split(":");
			if (strs.length != 2) {
				throw new IllegalArgumentException(s);
			}
			Point topLeft = ExcelUtils.nameToPoint(strs[0]);
			Point bottomRight = ExcelUtils.nameToPoint(strs[1]);
			this.readRange = new Rectangle(topLeft.x, topLeft.y, bottomRight.x, bottomRight.y);
		}
	}
	
	public CellValueHelper getCellValueHelper() { return this.helper;}
	
	private int getRowHeight(int row) {
		Row rowObj = sheet.getRow(row);
		if (rowObj != null && rowObj.getZeroHeight()) {
			return 0;
		}
		return ExcelUtils.getRowHeight(this.sheet, row);
	}
	
	private int getColumnWidth(int col) {
		if (this.sheet.isColumnHidden(col)) {
			return 0;
		}
		return this.columnWidth.getColumnWidth(this.sheet, col);
	}
	
	private void calcMaxCol(int target) {
		if (target < this.maxCol) {
			return;
		}
		if (this.readMaxCol > 0 && target > this.readMaxCol) {
			if (this.readMaxCol - 1 > this.maxCol) {
				target = this.readMaxCol - 1;
			} else {
				return;
			}
		}
		if (this.readRange != null && target > this.readRange.width) {
			if (this.readRange.width + 1 > this.maxCol) {
				target = this.readRange.width + 1;
			} else {
				return;
			}
		}
		this.maxCol = target;
	}
	
	public CellWrapper getCellWrapper(String id) {
		if (this.cells == null) {
			return null;
		}
		Point p = ExcelUtils.nameToPoint(id);
		int y = p.y - this.cells[0][0].getRow();
		int x = p.x - this.cells[0][0].getCol();
		if (y < cells.length && x < cells[y].length) {
			return cells[y][x];
		} else {
			return null;
		}
	}
	
	private ExcelToCanvas build() {
		ExcelToCanvas ret = new ExcelToCanvas();
		
		int startRow = 0;
		int startCol = 0;
		int width = 0;
		int height = 0;
		
		//行数と列数のカウント
		this.maxRow = sheet.getLastRowNum() + 1;
		if (this.readMaxRow > 0 && this.maxRow > this.readMaxRow) {
			this.maxRow = this.readMaxRow;
		}
		if (this.includeCells != null) {
			for (String cell : this.includeCells) {
				int idx = cell.lastIndexOf("!");
				if (idx != -1) {
					cell = cell.substring(idx + 1);
				}
				Point p = ExcelUtils.nameToPoint(cell);
				if (p.x > this.maxCol) {
					this.maxCol = p.x;
				}
				if (p.y > this.maxRow) {
					this.maxRow = p.y;
				}
			}
		}
		if (this.readRange != null) {
			startRow = this.readRange.y;
			startCol = this.readRange.x;
			if (this.maxRow > this.readRange.height + 1) {
				this.maxRow = this.readRange.height + 1;
			}
		}
		for (int i=startRow; i<this.maxRow; i++) {
			height += getRowHeight(i);
			Row row = sheet.getRow(i);
			if (row == null) {
				continue;
			}
			calcMaxCol(row.getLastCellNum());
		}
		for (int i=startCol; i<this.maxCol; i++) {
			width += getColumnWidth(i);
		}
		ret.setWidth(width);
		ret.setHeight(height);
		
		//CellWrapperの作成
		this.cells = new CellWrapper[this.maxRow - startRow][this.maxCol - startCol];
		int y = 0;
		for (int i=startRow; i<this.maxRow; i++) {
			Row row = sheet.getRow(i);
			int x = 0;
			for (int j=startCol; j<this.maxCol; j++) {
				Cell cell = row == null ? null : row.getCell(j);
				CellWrapper cw = new CellWrapper(sheet, cell, i, j, x, y);
				this.cells[i-startRow][j-startCol] = cw;
				x += getColumnWidth(j);
			}
			y += getRowHeight(i);
		}
		//結合セル情報
		for (int i=0; i<sheet.getNumMergedRegions(); i++) {
			CellRangeAddress a = sheet.getMergedRegion(i);
			for (int row=a.getFirstRow(); row<=a.getLastRow(); row++) {
				if (row >= startRow && row < this.maxRow) {
					for (int col=a.getFirstColumn(); col<=a.getLastColumn(); col++) {
						if (col >= startCol && col < this.maxCol) {
							this.cells[row-startRow][col-startCol].setMergedRegion(a);
						}
					}
				}
			}
		}
		//背景色
		//ToDo 隣り合うセルが同じ背景の場合は結合して最適化する
		for (int i=startRow; i<this.maxRow; i++) {
			Row row = sheet.getRow(i);
			int h = getRowHeight(i);
			for (int j=startCol; j<this.maxCol; j++) {
				int w = getColumnWidth(j);
				CellWrapper cw = this.cells[i-startRow][j-startCol];
				if (!this.includeHiddenCell && cw.isHidden()) {
					continue;
				}
				Cell cell = cw.getCell();
				if (cell != null) {
					CellStyle style = cell.getCellStyle();
					if (style != null) {
						String strBack = FormattedValue.getColorString(this.exColor.getFillBackgroundColor(cell));
						String strFore = FormattedValue.getColorString(this.exColor.getFillForegroundColor(cell));
						short pattern = style.getFillPattern();
						if (strBack != null ||
						    strFore != null ||
						    pattern != CellStyle.NO_FILL)
						{
							ret.addFillInfo(new ExcelToCanvas.FillInfo(cw.getLeft(), cw.getTop(), w, h, strBack, strFore, pattern));
							updateWriteCell(i, j);
						}
					}
				}
				ExcelToCanvas.StrInfo strInfo = cw.getStrInfo();
				if (strInfo != null) {
					ret.addStrInfo(strInfo);
				}
				if (cw.isWrite()) {
					updateWriteCell(i, j);
				}
			}
		}
		buildHorizontalLines(ret, startRow, this.maxRow, startCol, maxCol);
		buildVerticalLines(ret, startRow, this.maxRow, startCol, maxCol);
		if (this.includePicture) {
			buildPictures(ret);
		}
		if (this.includeChart) {
			buildCharts(ret);
		}
		
		//高さと幅の再計算
		if (this.writeRow + 1 < this.maxRow - 1) {
			height = 0;
			for (int i=startRow; i<=this.writeRow + 1; i++) {
				height += getRowHeight(i);
			}
			ret.setHeight(height);
		} else if (this.writeRow > this.maxRow - 1) {
			for (int i=this.maxRow; i<=this.writeRow; i++) {
				height += getRowHeight(i);
			}
			ret.setHeight(height);
		}
		if (this.writeCol + 1 < this.maxCol - 1) {
			width = 0;
			for (int i=startCol; i<=this.writeCol + 1; i++) {
				width += getColumnWidth(i);
			}
			ret.setWidth(width);
		} else if (this.writeCol > this.maxCol - 1) {
			for (int i=this.maxCol; i<=this.writeCol; i++) {
				width += getColumnWidth(i);
			}
			ret.setWidth(width);
		}
		//Debug
		/*
		{
			for (int i=0; i<this.maxRow; i++) {
				for (int j=0; j<this.maxCol; j++) {
					CellWrapper cw = cells[i][j];
					String text = cw.getText();
					if (text != null) {
						System.out.println("Text: " + text + ", " + cw.getCell().getCellStyle().getDataFormat() + ", " + cw.getCell().getCellStyle().getDataFormatString());
					}
				}
			}
		}
		*/
		return ret;
	}
	
	private void updateWriteCell(int row, int col) {
		if (row > this.writeRow) {
			this.writeRow = row;
		}
		if (col > this.writeCol) {
			this.writeCol = col;
		}
	}
	
	private int getColLeft(int col) {
		int ret = 0;
		for (int i=0; i<col; i++) {
			ret += getColumnWidth(i);
		}
		return ret;
	}
	
	private int getRowTop(int row) {
		int ret = 0;
		for (int i=0; i<row; i++) {
			ret += getRowHeight(i);
		}
		return ret;
	}
	
	private void buildPictures(ExcelToCanvas ret) {
		List<? extends PictureData> pictures = workbook.getAllPictures();
		if (pictures == null || pictures.size() == 0) {
			return;
		}
		try {
			if (this.workbook instanceof HSSFWorkbook) {
				HSSFPatriarch patriarch = ((HSSFSheet)this.sheet).getDrawingPatriarch();
				if (patriarch == null) {
					return;
				}
				List<HSSFShape> shapes = patriarch.getChildren();
				if (shapes == null) {
					return;
				}
				for (HSSFShape shape : shapes) {
					if (shape instanceof HSSFPicture) {
						HSSFPicture picture = (HSSFPicture)shape;
						/*
						*/
						HSSFAnchor anchor = picture.getAnchor();
						if (anchor instanceof HSSFClientAnchor) {
							HSSFClientAnchor ca = (HSSFClientAnchor)anchor;
							CellWrapper cw1 = this.cells[ca.getRow1()][ca.getCol1()];
							CellWrapper cw2 = this.cells[ca.getRow2()][ca.getCol2()];
							int[] p = new int[4];
							p[0] = cw1.getLeft() + (getColumnWidth(ca.getCol1()) * ca.getDx1() / 1024);
							p[1] = cw1.getTop() + (getRowHeight(ca.getRow1()) * ca.getDy1() / 256 );
							p[2] = (cw2.getLeft() + (getColumnWidth(ca.getCol2()) * ca.getDx2() / 1024)) - p[0];
							p[3] = (cw2.getTop() + (getRowHeight(ca.getRow2()) * ca.getDy2() / 256 )) - p[1];
							/*
							String border = null;
							int lineStyle = picture.getLineStyle();
							int lineColor = picture.getLineStyleColor();
							int lineWidth = picture.getLineWidth();
							System.out.println("test1: " + lineStyle + ", " + lineColor + ", " + lineWidth);
							*/
							addPictureInfo(ret, (PictureData)picture.getPictureData(), p, null);
						}
					}
				}
			} else if (this.workbook instanceof XSSFWorkbook) {
				XSSFDrawing drawing = ((XSSFSheet)this.sheet).createDrawingPatriarch();
				if (drawing == null) {
					return;
				}
				for (CTTwoCellAnchor anchor : drawing.getCTDrawing().getTwoCellAnchorArray()) {
					CTPicture picture = anchor.getPic();
					if (picture == null) {
						continue;
					}
					/*
					おそらくSpPr#getLnでボーダーが取得できる
					<a:ln w="12700">
					  <a:solidFill>
					  <a:srgbClr val="FF0000"/>
					  </a:solidFill>
					 <a:prstDash val="sysDash"/>
					</a:ln>
					*/
					/*
					ret.addPictureInfo(new ExcelToCanvas.PictureInfo((PictureData)picture.getPictureData(), p, null));
					*/
					int[] p = anchorToPoints(anchor);
					if (p == null) {
						continue;
					}
					
					PictureData data = null;
					String blipId = picture.getBlipFill().getBlip().getEmbed();
					for (POIXMLDocumentPart part : drawing.getRelations()) {
						if(part.getPackageRelationship().getId().equals(blipId)){
							data = (PictureData)part;
						}
					}
					/*
						p[0] = (int)(picture.getCTPicture().getSpPr().getXfrm().getOff().getX() / XSSFShape.EMU_PER_PIXEL);
						p[1] = (int)(picture.getCTPicture().getSpPr().getXfrm().getOff().getY() / XSSFShape.EMU_PER_PIXEL);
						p[2] = (int)(picture.getCTPicture().getSpPr().getXfrm().getExt().getCx() / XSSFShape.EMU_PER_PIXEL);
						p[3] = (int)(picture.getCTPicture().getSpPr().getXfrm().getExt().getCy() / XSSFShape.EMU_PER_PIXEL);
					*/
					addPictureInfo(ret, data, p, null);
				}
			} else {
				throw new IllegalArgumentException();
			}
		} catch (Exception e) {
			//サポート外の図形がある場合？
			e.printStackTrace();
		}
	}
	
	private int[] anchorToPoints(CTTwoCellAnchor anchor) {
		int sCol = anchor.getFrom().getCol();
		int eCol = anchor.getTo().getCol();
		int sRow = anchor.getFrom().getRow();
		int eRow = anchor.getTo().getRow();
		if (this.readRange != null) {
			if (!(sCol >= this.readRange.x && eCol <= this.readRange.width) ||
			    !(sRow >= this.readRange.y && eRow <= this.readRange.height)) {
				return null;
			}
		}
		updateWriteCell(eRow, eCol);
		
		int[] p = new int[4];
		p[0] = (int) (getColLeft(sCol) + (anchor.getFrom().getColOff() / XSSFShape.EMU_PER_PIXEL));
		p[1] = (int) (getRowTop(sRow) + (anchor.getFrom().getRowOff() / XSSFShape.EMU_PER_PIXEL));
		p[2] = (int) (getColLeft(eCol) + (anchor.getTo().getColOff() / XSSFShape.EMU_PER_PIXEL) - p[0]);
		p[3] = (int) (getRowTop(eRow) + (anchor.getTo().getRowOff() / XSSFShape.EMU_PER_PIXEL) - p[1]);
		return p;
	}
	
	private void addPictureInfo(ExcelToCanvas ret, PictureData data, int[] p, String border) {
		if (this.readRange != null) {
			int sx = getColLeft(this.readRange.x);
			int sy = getRowTop(this.readRange.y);
			p[0] -= sx;
			p[1] -= sy;
			if (p[0] < 0 || p[1] < 0) {
				return;
			}
			if (p[0] + p[2] > sx + ret.getWidth()  || p[1] + p[3] > sy + ret.getHeight()) {
				return;
			}
		}
		ret.addPictureInfo(new ExcelToCanvas.PictureInfo(data, p, border));
	}
	
	private void buildCharts(ExcelToCanvas ret) {
		if (!(this.sheet instanceof XSSFSheet)) {
			return;
		}
		XSSFDrawing drawing = ((XSSFSheet)this.sheet).createDrawingPatriarch();
		List<XSSFChart> chartList = drawing.getCharts();
		if (chartList == null || chartList.size() == 0) {
			return;
		}
		List<CTTwoCellAnchor> anchorList = new ArrayList<CTTwoCellAnchor>();
		for (CTTwoCellAnchor anchor : drawing.getCTDrawing().getTwoCellAnchorArray()) {
			if (anchor.isSetGraphicFrame()) {
				anchorList.add(anchor);
			}
		}
		if (chartList.size() != anchorList.size()) {
			throw new IllegalStateException();
		}
		XSSFWorkbook xWorkbook = (XSSFWorkbook)this.workbook;
		ChartFactory factory = getChartFactory();
		if (factory == null) {
			factory = new Flotr2ChartFactory();
			setChartFactory(factory);
		}
		factory.setIncludeRawData(this.includeRawData);
		for (int i=0; i<chartList.size(); i++) {
			XSSFChart xChart = chartList.get(i);
			CTTwoCellAnchor anchor = anchorList.get(i);
			int[] p = anchorToPoints(anchor);
			Chart chart = p == null ? null : factory.createChart(xWorkbook, xChart);
			if (p != null && chart != null) {
				ret.addChartInfo(new ExcelToCanvas.ChartInfo(p, chart));
			}
		}
	}
	
	private void buildHorizontalLines(ExcelToCanvas ret, int startRow, int maxRow, int startCol, int maxCol) {
		int h = 0;
		for (int i=startRow; i<maxRow + 1; i++) {
			int y = h;
			h += getRowHeight(i);
			
			Row row1 = i == startRow ? null : sheet.getRow(i-1);
			Row row2 = sheet.getRow(i);
			if (row1 == null && row2 == null) {
				continue;
			}
			
			int sx = -1;
			int kind = CellStyle.BORDER_NONE;
			String strColor = null;
			int x = 0;
			for (int j=startCol; j<maxCol; j++) {
				CellWrapper cw1 = i == startRow ? null : this.cells[i - startRow - 1][j - startCol];
				CellWrapper cw2 = i == maxRow ? null : this.cells[i - startRow][j - startCol];
				
				Cell cell1 = checkMergedRegion(cw1, 1);
				Cell cell2 = checkMergedRegion(cw2, 2);
				
				BorderInfo info = getBorder(cell1, cell2, true);
				if (info == null) {
					if (sx >= 0) {
						ret.addLineInfo(new ExcelToCanvas.LineInfo(sx, y, x, y, kind, strColor));
						updateWriteCell(i, j);
					}
					sx = -1;
					kind = CellStyle.BORDER_NONE;
					strColor = null;
				} else {
					if (sx < 0) {
						sx = x;
						kind = info.getKind();
						strColor = info.getColor();
					} else if (kind != info.getKind() || !Objects.equals(strColor, info.getColor())) {
						ret.addLineInfo(new ExcelToCanvas.LineInfo(sx, y, x, y, kind, strColor));
						updateWriteCell(i, j);
						sx = x;
						kind = info.getKind();
						strColor = info.getColor();
					}
				}
				x += getColumnWidth(j);
			}
			if (sx >= 0) {
				ret.addLineInfo(new ExcelToCanvas.LineInfo(sx, y, x, y, kind, strColor));
				updateWriteCell(i, maxCol - 1);
			}
		}
	}
	
	private void buildVerticalLines(ExcelToCanvas ret, int startRow, int maxRow, int startCol, int maxCol) {
		int w = 0;
		for (int i=startCol; i<maxCol + 1; i++) {
			int x = w;
			w += getColumnWidth(i);
			
			int sy = -1;
			int kind = CellStyle.BORDER_NONE;
			String strColor = null;
			int y = 0;
			for (int j=startRow; j<maxRow; j++) {
				CellWrapper cw1 = i == startCol ? null : this.cells[j - startRow][i - startCol - 1];
				CellWrapper cw2 = i == maxCol ? null : this.cells[j - startRow][i - startCol];
				
				Cell cell1 = checkMergedRegion(cw1, 3);
				Cell cell2 = checkMergedRegion(cw2, 4);
				
				BorderInfo info = getBorder(cell1, cell2, false);
				if (info == null) {
					if (sy >= 0) {
						ret.addLineInfo(new ExcelToCanvas.LineInfo(x, sy, x, y, kind, strColor));
						updateWriteCell(j, i);
					}
					sy = -1;
					kind = CellStyle.BORDER_NONE;
					strColor = null;
				} else {
					if (sy < 0) {
						sy = y;
						kind = info.getKind();
						strColor = info.getColor();
					} else if (kind != info.getKind() || !Objects.equals(strColor, info.getColor())) {
						ret.addLineInfo(new ExcelToCanvas.LineInfo(x, sy, x, y, kind, strColor));
						updateWriteCell(j, i);
						sy = y;
						kind = info.getKind();
						strColor = info.getColor();
					}
				}
				y += getRowHeight(j);
			}
			if (sy >= 0) {
				ret.addLineInfo(new ExcelToCanvas.LineInfo(x, sy, x, y, kind, strColor));
				updateWriteCell(maxRow-1, i);
			}
		}
	}
	
	private Cell checkMergedRegion(CellWrapper cw, int checkType) {
		if (cw == null || cw.getCell() == null) {
			return null;
		}
		boolean b = true;
		switch (checkType) {
			case 1:
				b = cw.isBottomSide();
				break;
			case 2:
				b = cw.isTopSide();
				break;
			case 3:
				b = cw.isRightSide();
				break;
			case 4:
				b = cw.isLeftSide();
				break;
			default:
				throw new IllegalStateException();
		}
		return b ? cw.getCell() : null;
	}
	
	private BorderInfo getBorder(Cell cell1, Cell cell2, boolean horizontal) {
		if (cell1 != null) {
			CellStyle style = cell1.getCellStyle();
			if (style != null) {
				int n = horizontal ? style.getBorderBottom() : style.getBorderRight();
				if (n != CellStyle.BORDER_NONE) {
					String strColor = FormattedValue.getColorString(horizontal ? exColor.getBottomBorderColor(cell1) : exColor.getRightBorderColor(cell1));
					return new BorderInfo(n, strColor);
				}
			}
		}
		if (cell2 != null) {
			CellStyle style = cell2.getCellStyle();
			if (style != null) {
				int n = horizontal ? style.getBorderTop() : style.getBorderLeft();
				if (n != CellStyle.BORDER_NONE) {
					String strColor = FormattedValue.getColorString(horizontal ? exColor.getTopBorderColor(cell2) : exColor.getLeftBorderColor(cell2));
					return new BorderInfo(n, strColor);
				}
			}
		}
		return null;
	}
	
	private static String getAlignString(int align, int valign, boolean general) {
		StringBuilder buf = new StringBuilder();
		switch (align) {
			case CellStyle.ALIGN_CENTER:
			case CellStyle.ALIGN_CENTER_SELECTION:
			case CellStyle.ALIGN_FILL:
			case 7://均等割り付け
				buf.append("c");
				break;
			case CellStyle.ALIGN_LEFT:
				buf.append("l");
				break;
			case CellStyle.ALIGN_JUSTIFY:
				buf.append("j");
				break;
			case CellStyle.ALIGN_RIGHT:
				buf.append("r");
				break;
			case CellStyle.ALIGN_GENERAL:
			default:
				throw new IllegalStateException(Integer.toString(align));
		}
		switch (valign) {
			case CellStyle.VERTICAL_TOP:
				buf.append("t");
				break;
			case CellStyle.VERTICAL_CENTER:
				buf.append("c");
				break;
			case CellStyle.VERTICAL_BOTTOM:
				buf.append("b");
				break;
			case CellStyle.VERTICAL_JUSTIFY:
				buf.append("j");
				break;
			default:
				//valignは未設定がありえる
				buf.append("c");
				break;
		}
		if (general) {
			buf.append("g");
		}
		return buf.toString();
	}
	
	public class CellWrapper {
		
		private Sheet sheet;
		private Cell cell;
		private int col;
		private int row;
		private int left;
		private int top;
		private CellRangeAddress mergedRegion;
		private FormattedValue formattedValue;
		
		public CellWrapper(Sheet sheet, Cell cell, int row, int col, int left, int top) {
			this.sheet = sheet;
			this.cell = cell;
			this.row = row;
			this.col = col;
			this.left = left;
			this.top = top;
			if (cell != null) {
				cell.setCellStyle(getCellStyle());
			}
		}
		
		public String getId() {
			return ExcelUtils.pointToName(this.col, this.row);
		}
		
		public boolean isHidden() {
			Row rowObj = sheet.getRow(this.row);
			if (rowObj != null && rowObj.getZeroHeight()) {
				return true;
			}
			return sheet.isColumnHidden(this.col);
		}
		
		public Cell getCell() { return this.cell;}
		
		public boolean isWrite() {
			if (this.mergedRegion != null) {
				return true;
			}
			return getText() != null;
		}
		
		public int getRow() { return this.row;}
		public int getCol() { return this.col;}
		public int getLeft() { return this.left;}
		public int getTop() { return this.top;}
		
		public CellRangeAddress getMergedRegion() { return this.mergedRegion;}
		public void setMergedRegion(CellRangeAddress a) { this.mergedRegion = a;}
		
		public boolean isMerged() { return this.mergedRegion != null;}
		
		public boolean isMainCell() {
			if (this.mergedRegion == null) {
				return true;
			}
			return this.mergedRegion.getFirstRow() == this.row && 
				this.mergedRegion.getFirstColumn() == this.col;
		}
		
		public boolean isLeftSide() {
			if (this.mergedRegion == null) {
				return true;
			}
			return this.mergedRegion.getFirstColumn() == this.col;
		}
		
		public boolean isRightSide() {
			if (this.mergedRegion == null) {
				return true;
			}
			return this.mergedRegion.getLastColumn() == this.col;
		}
		
		public boolean isTopSide() {
			if (this.mergedRegion == null) {
				return true;
			}
			return this.mergedRegion.getFirstRow() == this.row;
		}
		
		public boolean isBottomSide() {
			if (this.mergedRegion == null) {
				return true;
			}
			return this.mergedRegion.getLastRow() == this.row;
		}
		
		public int getHeight() {
			if (!isMerged()) {
				return getRowHeight(this.row);
			}
			if (!isMainCell()) {
				return 0;
			}
			int ret = 0;
			for (int i=this.mergedRegion.getFirstRow(); i<=this.mergedRegion.getLastRow(); i++) {
				ret += getRowHeight(i);
			}
			return ret;
		}
		
		public int getWidth() {
			if (!isMerged()) {
				return getColumnWidth(this.col);
			}
			if (!isMainCell()) {
				return 0;
			}
			int ret = 0;
			for (int i=this.mergedRegion.getFirstColumn(); i<=this.mergedRegion.getLastColumn(); i++) {
				ret += getColumnWidth(i);
			}
			return ret;
		}
		
		public ExcelToCanvas.StrInfo getStrInfo() {
			if (!isMainCell()) {
				return null;
			}
			String id = ExcelUtils.pointToName(col, row);
			String text = getText();
			boolean bIncludeCell = isIncludeCell(id);
			if (text == null) {
				if (includeEmptyStr || 
				    bIncludeCell ||
				    (includeRawData && isFormula()) ||
				    (includeComment && this.cell != null && this.cell.getCellComment() != null)
				   ) 
				{
					text = "";
				} else {
					return null;
				}
			}
			
			//Style
			boolean alignGeneral = true;
			int align = CellStyle.ALIGN_LEFT;
			int valign = CellStyle.VERTICAL_CENTER;
			
			Map<String, String> styleMap = new HashMap<String, String>();
			CellStyle style = getCellStyle();
			if (style != null) {
				align = style.getAlignment();
				if (align == CellStyle.ALIGN_GENERAL) {
					int type = cell == null ? Cell.CELL_TYPE_STRING : cell.getCellType();
					if (type == Cell.CELL_TYPE_FORMULA) {
						type = cell.getCachedFormulaResultType();
					}
					if (type == Cell.CELL_TYPE_NUMERIC || ExcelUtils.isNumericStyle(style)) {
						align = CellStyle.ALIGN_RIGHT;
					} else {
						align = CellStyle.ALIGN_LEFT;
					}
				} else {
					alignGeneral = false;
				}
				int indent = style.getIndention();
				switch (align) {
					case CellStyle.ALIGN_CENTER:
					case CellStyle.ALIGN_CENTER_SELECTION:
					case CellStyle.ALIGN_FILL:
					case 7://均等割り付け
						if (indent != 0) {
							styleMap.put("margin", "0 " + indent + "em");
						}
						break;
					case CellStyle.ALIGN_LEFT:
						if (indent != 0) {
							styleMap.put("margin-left", indent + "em");
						}
						break;
					case CellStyle.ALIGN_JUSTIFY:
						if (indent != 0) {
							styleMap.put("margin", "0 " + indent + "em");
						}
						break;
					case CellStyle.ALIGN_RIGHT:
						if (indent != 0) {
							styleMap.put("margin-right", "0 " + indent + "em");
						}
						break;
					case CellStyle.ALIGN_GENERAL:
					default:
						throw new IllegalStateException(Integer.toString(align));
				}
				
				valign = style.getVerticalAlignment();
				if (style.getRotation() != 0) {
					styleMap.put("transform", "rotate(" + style.getRotation() + ")");
				}
				if (!style.getWrapText()) {
					styleMap.put("text-wrap", "none");
				}
				Font font = workbook.getFontAt(style.getFontIndex());
				styleMap.put("font-family", fontManager.getFontFamily(font.getFontName()));
				styleMap.put("font-size", font.getFontHeightInPoints() + "pt");
				String color = FormattedValue.getColorString(exColor.getFontColor(font));
				if (this.formattedValue != null && this.formattedValue.getColor() != null) {
					color = FormattedValue.getColorString(this.formattedValue.getColor());
				}
				if (color != null) {
					styleMap.put("color", color);
				}
				if (font.getBoldweight() == Font.BOLDWEIGHT_BOLD) {
					styleMap.put("font-weight", "bold");
				}
				if (font.getItalic()) {
					styleMap.put("font-style", "italic");
				}
				byte ul = font.getUnderline();
				if (ul != Font.U_NONE) {
					styleMap.put("text-decoration", "underline");
					if (ul == Font.U_DOUBLE || ul == Font.U_DOUBLE_ACCOUNTING) {
						styleMap.put("text-decoration-style", "double");
					}
				} else if (font.getStrikeout()) {
					styleMap.put("text-decoration", "line-through");
				}
			}
			//Comment
			String comment = null;
			Integer commentWidth = null;
			if (includeComment) {
				comment = getComment();
				if (comment != null) {
					commentWidth = Integer.valueOf(getWidth());
				}
			}
			//座標
			int startRow = readRange == null ? 0 : readRange.y;
			int startCol = readRange == null ? 0 : readRange.x;
			int width = getWidth();
			if (comment == null && !bIncludeCell) {
				if (isExpandCellRight(style, align)) {
					for (int i=this.col + 1; i<maxCol; i++) {
						CellWrapper next = cells[this.row - startRow][i - startCol];
						CellStyle nextStyle = next.getCellStyle();
						if (!isIncludeCell(next.getId()) && next.getText() == null && (nextStyle == null || nextStyle.getBorderLeft() == CellStyle.BORDER_NONE)) {
							width += getColumnWidth(i);
							if (nextStyle != null && nextStyle.getBorderRight() != CellStyle.BORDER_NONE) {
								break;
							}
						} else {
							break;
						}
					}
				}
				if (isExpandCellLeft(style, align)) {
					for (int i=this.col - 1; i>=0; i--) {
						CellWrapper prev = cells[this.row - startRow][i - startCol];
						CellStyle prevStyle = prev.getCellStyle();
						if (!isIncludeCell(prev.getId()) && prev.getText() == null && (prevStyle == null || prevStyle.getBorderRight() == CellStyle.BORDER_NONE)) {
							int w = getColumnWidth(i);
							width += w;
							this.left -= w;
							if (prevStyle != null && prevStyle.getBorderLeft() != CellStyle.BORDER_NONE) {
								break;
							}
						} else {
							break;
						}
					}
				}
			}
			int[] p = new int[4];
			p[0] = getLeft();
			p[1] = getTop();
			p[2] = width;
			p[3] = getHeight();
			
			//リンク
			String link = null;
			if (this.cell != null && this.cell.getHyperlink() != null) {
				Hyperlink h = this.cell.getHyperlink();
				if (h.getType() == Hyperlink.LINK_EMAIL || h.getType() == Hyperlink.LINK_URL) {
					link = h.getAddress();
				}
			}
			//Formula
			boolean formula = this.cell != null && cell.getCellType() == Cell.CELL_TYPE_FORMULA;
			ExcelToCanvas.StrInfo ret = new ExcelToCanvas.StrInfo(p, id, text, getAlignString(align, valign, alignGeneral), styleMap, link, comment, commentWidth, formula);
			if (this.formattedValue != null && ExcelToCanvasBuilder.this.includeRawData) {
				FormattedValue.Type type = this.formattedValue.getType();
				if (isFormula()) {
					ret.setRawData(this.formattedValue.getFormula());
				} else if (type == FormattedValue.Type.NUMBER || type == FormattedValue.Type.DATE) {
					ret.setRawData(this.formattedValue.getRawString());
				}
			}
			return ret;
		}
		
		private CellStyle getCellStyle() {
			CellStyle style = this.cell == null ? null : this.cell.getCellStyle();
			if (style != null) {
				return style;
			}
			style = sheet.getColumnStyle(this.col);
			if (style != null) {
				return style;
			}
			Row row = sheet.getRow(this.row);
			if (row != null) {
				try {
					style = row.getRowStyle();
				} catch (IndexOutOfBoundsException e) {
					//Maybe HSSF bug
					System.err.println("Maybe HSSF bug");
					e.printStackTrace();
				}
			}
			return style;
		}
		
		private String getComment() {
			if (this.cell == null || this.cell.getCellComment() == null) {
				return null;
			}
			Comment comment = this.cell.getCellComment();
			String value = comment.getString() == null ? null : comment.getString().getString();
			if (value != null) {
				value = convertHtml(value);
			}
			return value;
		}
		
		public String getText() {
			if (this.cell == null || !isMainCell()) {
				return null;
			}
			if (this.formattedValue == null) {
				this.formattedValue = helper.getFormattedValue(this.cell);
				String value = this.formattedValue.getValue();
				if (value == null || value.length() == 0) {
					this.formattedValue.setValue(null);
					return null;
				}
				this.formattedValue.setValue(convertHtml(value));
			}
			return this.formattedValue.getValue();
		}
		
		private boolean isFormula() {
			return this.formattedValue != null && this.formattedValue.getFormula() != null;
		}
		
		private boolean isIncludeCell(String id) {
			if (includeCells == null) {
				return false;
			}
			return includeCells.contains(id) || includeCells.contains(sheet.getSheetName() + "!" + id);
		}

		private String convertHtml(String value) {
			StringBuilder buf = new StringBuilder();
			int len = value.length();
			for (int i=0; i<len; i++) {
				char c = value.charAt(i);
				switch (c) {
					case '<':
						buf.append("&lt;");
						break;
					case '>':
						buf.append("&gt;");
						break;
					case '"':
						buf.append("&quot;");
						break;
					case '&':
						buf.append("&amp;");
						break;
					case '\n':
						buf.append("<br>");
						break;
					default:
						buf.append(c);
				}
			}
			return buf.toString();
		}
		
		public String getColorString() {
			Color color = null;
			if (this.formattedValue != null) {
				color = this.formattedValue.getColor();
			}
			if (color == null && this.cell != null) {
				color = exColor.getFontColor(this.cell);
			}
			return color == null ? null : FormattedValue.getColorString(color);
		}
		
		private boolean isExpandCellRight(CellStyle style, int align) {
			if (ExcelToCanvasBuilder.this.includeEmptyStr) {
				return false;
			}
			if (align != CellStyle.ALIGN_LEFT) {
				return false;
			}
			if (this.mergedRegion != null) {
				return false;
			}
			if (style != null && style.getBorderRight() != CellStyle.BORDER_NONE) {
				return false;
			}
			if (ExcelToCanvasBuilder.this.expandChecker != null) {
				return ExcelToCanvasBuilder.this.expandChecker.isExpandCellRight(this);
			}
			return true;
		}

		private boolean isExpandCellLeft(CellStyle style, int align) {
			if (ExcelToCanvasBuilder.this.includeEmptyStr) {
				return false;
			}
			if (align != CellStyle.ALIGN_RIGHT) {
				return false;
			}
			if (this.mergedRegion != null) {
				return false;
			}
			if (style != null && style.getBorderLeft() != CellStyle.BORDER_NONE) {
				return false;
			}
			if (ExcelToCanvasBuilder.this.expandChecker != null) {
				return ExcelToCanvasBuilder.this.expandChecker.isExpandCellLeft(this);
			}
			return true;
		}
	}
	
	public interface ExpandChecker {
		
		public boolean isExpandCellRight(CellWrapper cw);
		public boolean isExpandCellLeft(CellWrapper cw);
		
	}
	
	private static class BorderInfo {
		
		private int kind;
		private String color;
		
		public BorderInfo(int k, String c) {
			this.kind = k;
			this.color = c;
		}
		
		public int getKind() { return this.kind;}
		public String getColor() { return this.color;}
	}

}
