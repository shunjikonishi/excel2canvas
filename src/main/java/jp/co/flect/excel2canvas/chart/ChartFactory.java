package jp.co.flect.excel2canvas.chart;

import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Factory interface of chart.
 */
public interface ChartFactory {
	
	public String getChartName();
	public Chart createChart(XSSFWorkbook workbook, XSSFChart chart);
	public boolean isIncludeRawData();
	public void setIncludeRawData(boolean b);
	
}
