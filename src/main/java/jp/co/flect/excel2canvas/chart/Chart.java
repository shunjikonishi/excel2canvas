package jp.co.flect.excel2canvas.chart;

import java.util.List;

/**
 * Chart interface
 */
public interface Chart {
	
	public List<String> getCellNames();
	public boolean setCellValue(String name, Object value);
	public void clearRawData();
	public Chart cloneChart();
}
