package jp.co.flect.excel2canvas;

import java.io.File;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.fail;
import org.junit.Test;

/**
 * Unit test for simple App.
 */
public class JsonTest {
	
	@Test
	public void compareFromTo() throws Exception {
		File f = new File("testdata/jsontest.xlsx");
		
		ExcelToCanvasBuilder builder = new ExcelToCanvasBuilder();
		builder.setIncludeComment(true);
		//builder.setIncludeChart(true);
		ExcelToCanvas excel1 = builder.build(f);
		String json1 = excel1.toJson();
		
		ExcelToCanvas excel2 = ExcelToCanvas.fromJson(json1);
		String json2 = excel2.toJson();
		
		assertEquals(json1, json2);
	}
}
