package jp.co.flect.excel2canvas;

import java.io.File;
import java.util.List;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.fail;
import org.junit.Test;

import jp.co.flect.excel2canvas.ExcelToCanvas.StrInfo;

public class ChartTest {

	@Test
	public void parse() throws Exception {
		File f = new File("testdata/graph.xlsx");
		ExcelToCanvasBuilder builder = new ExcelToCanvasBuilder();
		builder.setIncludeChart(true);
		ExcelToCanvas excel = builder.build(f);

	}
}