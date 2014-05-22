package jp.co.flect.excel2canvas.chart;

import java.util.ArrayList;
import java.util.List;
import org.w3c.dom.NodeList;
import org.w3c.dom.Text;

import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTRadarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBubbleChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBubbleSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTValAx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTScaling;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarDir;
import org.openxmlformats.schemas.drawingml.x2006.chart.STBarGrouping;
import org.openxmlformats.schemas.drawingml.x2006.chart.STGrouping;
import org.openxmlformats.schemas.drawingml.x2006.chart.STRadarStyle;

import jp.co.flect.excel2canvas.chart.Flotr2.NameInfo;
import jp.co.flect.excel2canvas.ExcelUtils;

/**
 * ChartFactory implementation by Flotr2
 */
public class Flotr2ChartFactory implements ChartFactory {
	
	private static final String NAMESPACE_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
	private static final String NAMESPACE_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    
    private boolean includeRawData;
    private List<NameInfo[]> cellNames = new ArrayList<NameInfo[]>();
    
	public boolean isIncludeRawData() { return this.includeRawData;}
	public void setIncludeRawData(boolean b) { this.includeRawData = b;}
	
	public String getChartName() { return "Flotr2";}
	
	private void addNameInfo(NameInfo n) {
		if (n == null) {
			return;
		}
		NameInfo[] arr = new NameInfo[1];
		arr[0] = n;
		this.cellNames.add(arr);
	}
	
	private void addNameInfo(NameInfo[] ns) {
		if (ns == null || ns.length == 0) {
			return;
		}
		this.cellNames.add(ns);
	}
	
	public Chart createChart(XSSFWorkbook workbook, XSSFChart xc) {
		this.cellNames.clear();
		CTChart ctChart = xc.getCTChart();
		if (ctChart == null) {
			return null;
		}
		CTPlotArea plotArea = ctChart.getPlotArea();
		if (plotArea == null) {
			return null;
		}
		if (plotArea.sizeOfPieChartArray() > 0) {
			return createPieChart(workbook, ctChart, plotArea.getPieChartArray(0));
		}
		if (plotArea.sizeOfBarChartArray() > 0) {
			return createBarChart(workbook, ctChart, plotArea.getBarChartArray(0));
		}
		if (plotArea.sizeOfLineChartArray() > 0) {
			return createLineChart(workbook, ctChart, plotArea.getLineChartArray(0));
		}
		if (plotArea.sizeOfRadarChartArray() > 0) {
			return createRadarChart(workbook, ctChart, plotArea.getRadarChartArray(0));
		}
		if (plotArea.sizeOfBubbleChartArray() > 0) {
			return createBubbleChart(workbook, ctChart, plotArea.getBubbleChartArray(0));
		}
		return null;
	}
	
	private Chart createPieChart(XSSFWorkbook workbook, CTChart ctChart, CTPieChart pieChart) {
		Flotr2 ret = new Flotr2(Flotr2.Type.PIE);
		for (int i=0; i<pieChart.sizeOfSerArray(); i++) {
			processSeries(ret, workbook, new PieSeries(pieChart.getSerArray(i)));
		}
		commonSetting(ret, ctChart);
		return ret;
	}
	
	private Chart createBarChart(XSSFWorkbook workbook, CTChart ctChart, CTBarChart barChart) {
		Flotr2 ret = new Flotr2(Flotr2.Type.BAR);
		ret.setSeriesCount(barChart.sizeOfSerArray());
		ret.getOption().bars.horizontal = barChart.getBarDir().getVal() == STBarDir.BAR;
		ret.getOption().bars.stacked = barChart.getGrouping().getVal() == STBarGrouping.STACKED;
		if (!ret.getOption().bars.stacked) {
			ret.getOption().bars.barWidth = 1.0;
		}
		for (int i=0; i<barChart.sizeOfSerArray(); i++) {
			processSeries(ret, workbook, new BarSeries(barChart.getSerArray(i)));
		}
		commonSetting(ret, ctChart);
		return ret;
	}
	
	private Chart createLineChart(XSSFWorkbook workbook, CTChart ctChart, CTLineChart lineChart) {
		Flotr2 ret = new Flotr2(Flotr2.Type.LINE);
		ret.getOption().lines.stacked = lineChart.getGrouping().getVal() == STGrouping.STACKED;
		for (int i=0; i<lineChart.sizeOfSerArray(); i++) {
			processSeries(ret, workbook, new LineSeries(lineChart.getSerArray(i)));
		}
		commonSetting(ret, ctChart);
		return ret;
	}
	
	private Chart createRadarChart(XSSFWorkbook workbook, CTChart ctChart, CTRadarChart radarChart) {
		Flotr2 ret = new Flotr2(Flotr2.Type.RADAR);
		ret.getOption().radar.fill = radarChart.getRadarStyle().getVal() == STRadarStyle.FILLED;
		for (int i=0; i<radarChart.sizeOfSerArray(); i++) {
			processSeries(ret, workbook, new RadarSeries(radarChart.getSerArray(i)));
		}
		commonSetting(ret, ctChart);
		return ret;
	}
	
	private Chart createBubbleChart(XSSFWorkbook workbook, CTChart ctChart, CTBubbleChart bubbleChart) {
		Flotr2 ret = new Flotr2(Flotr2.Type.BUBBLE);
		for (int i=0; i<bubbleChart.sizeOfSerArray(); i++) {
			BubbleSeries series = new BubbleSeries(bubbleChart.getSerArray(i));
			addNameInfo(series.getNameInfo());
			
			String seriesName = series.getName();
			List<Double> xList = getValues(workbook, series.getCat());
			List<Double> yList = getValues(workbook, series.getVal());
			List<Double> sizeList = getValues(workbook, series.getSize());
			if (xList.size() == yList.size() && xList.size() == sizeList.size()) {
				for (int j=0; j<xList.size(); j++) {
					ret.addBubble(seriesName, xList.get(j), yList.get(j), sizeList.get(j));
				}
			}
		}
		if (bubbleChart.isSetBubbleScale()) {
			ret.getOption().bubbles = new Flotr2.BubbleOption(bubbleChart.getBubbleScale().getVal());
		}
		commonSetting(ret, ctChart);
		return ret;
	}
	
	private void commonSetting(Flotr2 ret, CTChart ctChart) {
		ret.getOption().title = getTitle(ctChart);
		if (ctChart.getPlotArea().sizeOfValAxArray() > 0) {
			CTValAx valAx = ctChart.getPlotArea().getValAxArray(0);
			CTScaling scaling = valAx.getScaling();
			if (scaling != null && (scaling.isSetMin() || scaling.isSetMax())) {
				boolean horizontal = ret.getOption().bars != null && ret.getOption().bars.horizontal;
				Flotr2.Axis axis = horizontal ? ret.getOption().xaxis : ret.getOption().yaxis;
				if (axis == null) {
					axis = new Flotr2.Axis();
					if (horizontal) {
						ret.getOption().xaxis = axis;
					} else {
						ret.getOption().yaxis = axis;
					}
				}
				if (scaling.isSetMin()) {
					axis.min = scaling.getMin().getVal();
				}
				if (scaling.isSetMax()) {
					axis.max = scaling.getMax().getVal();
				}
			}
		}
		if (this.includeRawData) {
			ret.setCellNames(this.cellNames);
		}
	}
	
	private void processSeries(Flotr2 ret, XSSFWorkbook workbook, SeriesWrapper series) {
		addNameInfo(series.getNameInfo());
		String seriesName = series.getName();
		if (series.isCatNumber()) {
			List<Double> xValues = getValues(workbook, series.getCat());
			List<Double> yValues = getValues(workbook, series.getVal());
			if (xValues.size() != yValues.size()) {
				while (xValues.size() < yValues.size()) {
					xValues.add(0.0);
				}
				while (yValues.size() < xValues.size()) {
					yValues.add(0.0);
				}
			}
			int len = xValues.size();
			for (int i=0; i<len; i++) {
				ret.addData(seriesName, xValues.get(i), yValues.get(i));
			}
		} else {
			List<String> names = getNames(workbook, series.getCat());
			List<Double> values = getValues(workbook, series.getVal());
			if (names.size() != values.size()) {
				while (names.size() < values.size()) {
					names.add("");
				}
				while (values.size() < names.size()) {
					values.add(0.0);
				}
			}
			int len = names.size();
			for (int i=len-1; i>=0; i--) {
				String name = names.get(i);
				double d = values.get(i);
				if ("".equals(name) && d == 0.0) {
					len--;
				} else {
					break;
				}
			}
			for (int i=0; i<len; i++) {
				ret.addData(seriesName, names.get(i), values.get(i));
			}
		}
	}
	
	private List<String> getNames(XSSFWorkbook workbook, CTAxDataSource src) {
		if (src != null && src.isSetStrRef()) {
			return getNames(workbook, src.getStrRef());
		}
		return new ArrayList<String>();
	}
	
	private List<String> getNames(XSSFWorkbook workbook, CTStrRef strRef) {
		List<String> ret = new ArrayList<String>();
		try {
			CellReference[] cells = new AreaReference(strRef.getF()).getAllReferencedCells();
			Sheet sheet = workbook.getSheet(cells[0].getSheetName());
			if (sheet == null) {
				throw new Exception(strRef.getF());
			}
			int idx = 0;
			NameInfo[] names = new NameInfo[cells.length];
			for (CellReference ref : cells) {
				names[idx++] = new NameInfo(NameInfo.TYPE_NAME, ExcelUtils.pointToName(ref.getCol(), ref.getRow()));
				String str = null;
				Row row = sheet.getRow(ref.getRow());
				if (row != null) {
					Cell cell = row.getCell(ref.getCol());
					if (cell != null) {
						str = cell.getStringCellValue();
					}
				}
				if (str == null) {
					str = "";
				}
				ret.add(str);
			}
			addNameInfo(names);
		} catch (Exception e) {
			e.printStackTrace();
			ret.clear();
			if (strRef.isSetStrCache()) {
				CTStrData data = strRef.getStrCache();
				for (int i=0; i<data.sizeOfPtArray(); i++) {
					CTStrVal pt = data.getPtArray(i);
					ret.add(pt.getV());
				}
			}
		}
		return ret;
	}
	
	private List<Double> getValues(XSSFWorkbook workbook, CTAxDataSource src) {
		if (src.isSetNumRef()) {
			return getValues(workbook, src.getNumRef());
		}
		return new ArrayList<Double>();
	}
	
	private List<Double> getValues(XSSFWorkbook workbook, CTNumDataSource src) {
		if (src.isSetNumRef()) {
			return getValues(workbook, src.getNumRef());
		}
		return new ArrayList<Double>();
	}
	
	private List<Double> getValues(XSSFWorkbook workbook, CTNumRef numRef) {
		List<Double> ret = new ArrayList<Double>();
		try {
			CellReference[] cells = new AreaReference(numRef.getF()).getAllReferencedCells();
			Sheet sheet = workbook.getSheet(cells[0].getSheetName());
			if (sheet == null) {
				throw new Exception(numRef.getF());
			}
			int idx = 0;
			NameInfo[] names = new NameInfo[cells.length];
			for (CellReference ref : cells) {
				names[idx++] = new NameInfo(NameInfo.TYPE_VALUE, ExcelUtils.pointToName(ref.getCol(), ref.getRow()));
				double d = 0.0;
				Row row = sheet.getRow(ref.getRow());
				if (row != null) {
					Cell cell = row.getCell(ref.getCol());
					if (cell != null) {
						d = cell.getNumericCellValue();
					}
				}
				ret.add(d);
			}
			addNameInfo(names);
		} catch (Exception e) {
			e.printStackTrace();
			ret.clear();
			if (numRef.isSetNumCache()) {
				CTNumData data = numRef.getNumCache();
				for (int i=0; i<data.sizeOfPtArray(); i++) {
					CTNumVal pt = data.getPtArray(i);
					ret.add(Double.parseDouble(pt.getV()));
				}
			}
		}
		return ret;
	}
	
	private static String getTitle(CTChart ctChart) {
		if(!ctChart.isSetTitle()) {
			return null;
		}
		return getString(ctChart.getTitle(), "declare namespace a='" + NAMESPACE_A + "' .//a:t");
	}
	
	private static String getString(XmlObject obj, String path) {
		XmlObject[] t = obj.selectPath(path);
		if (t == null || t.length == 0) {
			return null;
		}
		StringBuilder buf = new StringBuilder();
		for (int m = 0; m < t.length; m++) {
			NodeList kids = t[m].getDomNode().getChildNodes();
			for (int n = 0; n < kids.getLength(); n++) {
				if (kids.item(n) instanceof Text) {
					buf.append(kids.item(n).getNodeValue());
				}
			}
		}
		return buf.length() > 0 ? buf.toString() : null;
	}
	
	private static abstract class SeriesWrapper {
		
		public abstract boolean isSetTx();
		public abstract CTSerTx getTx();
		public abstract CTAxDataSource getCat();
		public abstract CTNumDataSource getVal();
		
		public NameInfo getNameInfo() {
			if (isSetTx()) {
				return new NameInfo(NameInfo.TYPE_TITLE, getString(getTx(), "declare namespace c='" + NAMESPACE_C + "' .//c:f"));
			}
			return null;
		}
		
		public String getName() {
			if (isSetTx()) {
				return getString(getTx(), "declare namespace c='" + NAMESPACE_C + "' .//c:v");
			}
			return null;
		}
		
		public boolean isCatNumber() {
			return getCat() == null ? false : getCat().isSetNumRef();
		}
	}
	
	private static class PieSeries extends SeriesWrapper {
		
		private CTPieSer series;
		
		public PieSeries(CTPieSer series) {
			this.series = series;
		}
		
		public boolean isSetTx() { return series.isSetTx();}
		public CTSerTx getTx() { return series.getTx();}
		public CTAxDataSource getCat() { return series.getCat();}
		public CTNumDataSource getVal() { return series.getVal();}
		
	}
	
	private static class BarSeries extends SeriesWrapper {
		
		private CTBarSer series;
		
		public BarSeries(CTBarSer series) {
			this.series = series;
		}
		
		public boolean isSetTx() { return series.isSetTx();}
		public CTSerTx getTx() { return series.getTx();}
		public CTAxDataSource getCat() { return series.getCat();}
		public CTNumDataSource getVal() { return series.getVal();}
		
	}
	
	private static class LineSeries extends SeriesWrapper {
		
		private CTLineSer series;
		
		public LineSeries(CTLineSer series) {
			this.series = series;
		}
		
		public boolean isSetTx() { return series.isSetTx();}
		public CTSerTx getTx() { return series.getTx();}
		public CTAxDataSource getCat() { return series.getCat();}
		public CTNumDataSource getVal() { return series.getVal();}
		
	}
	
	private static class RadarSeries extends SeriesWrapper {
		
		private CTRadarSer series;
		
		public RadarSeries(CTRadarSer series) {
			this.series = series;
		}
		
		public boolean isSetTx() { return series.isSetTx();}
		public CTSerTx getTx() { return series.getTx();}
		public CTAxDataSource getCat() { return series.getCat();}
		public CTNumDataSource getVal() { return series.getVal();}
		
	}
	
	private static class BubbleSeries extends SeriesWrapper {
		
		private CTBubbleSer series;
		
		public BubbleSeries(CTBubbleSer series) {
			this.series = series;
		}
		
		public boolean isSetTx() { return series.isSetTx();}
		public CTSerTx getTx() { return series.getTx();}
		public CTAxDataSource getCat() { return series.getXVal();}
		public CTNumDataSource getVal() { return series.getYVal();}
		
		public CTNumDataSource getSize() { return series.getBubbleSize();}
		
	}
	
}
