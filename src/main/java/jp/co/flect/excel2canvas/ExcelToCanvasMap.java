package jp.co.flect.excel2canvas;

import java.util.LinkedHashMap;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import jp.co.flect.excel2canvas.chart.Chart;
import jp.co.flect.excel2canvas.chart.Flotr2;

public class ExcelToCanvasMap extends LinkedHashMap<String, ExcelToCanvas> {

	public String toJson() {
		return toJson(false);
	}

	public String toJson(boolean indent) {
		GsonBuilder builder = new GsonBuilder();
		if (indent) {
			builder = builder.setPrettyPrinting();
		}
		return builder.create().toJson(this);
	}


	public static ExcelToCanvasMap fromJson(String json) {
		GsonBuilder builder = new GsonBuilder();
		builder.registerTypeAdapter(Chart.class, new Flotr2());
		ExcelToCanvasMap ret = builder.create().fromJson(json, ExcelToCanvasMap.class);
		for (ExcelToCanvas excel: ret.values()) {
			excel.loadInit();
		}
		return ret;
	}
}