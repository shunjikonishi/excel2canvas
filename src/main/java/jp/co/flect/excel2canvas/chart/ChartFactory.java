package jp.co.flect.excel2canvas.chart;

import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Factory interface of chart.
 */
public interface ChartFactory {
	
	public Chart createChart(XSSFWorkbook workbook, XSSFChart chart);
}
