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

/**
 * Unit test for simple App.
 */
public class InvisibleTest {
	
	@Test
	/**
	 * B列が非表示列のシートを読み込むテスト
	 */
	public void compareFromTo() throws Exception {
		File f = new File("testdata/InvisibleCol.xlsx");
		
		ExcelToCanvasBuilder builder = new ExcelToCanvasBuilder();
		builder.addIncludeCell("A2");
		builder.addIncludeCell("B2");
		builder.addIncludeCell("C2");
		builder.addIncludeCell("A3");
		builder.addIncludeCell("B3");
		builder.addIncludeCell("C3");
		builder.addIncludeCell("A4");
		builder.addIncludeCell("B4");
		builder.addIncludeCell("C4");
		builder.addIncludeCell("A5");
		builder.addIncludeCell("B5");
		builder.addIncludeCell("C5");
		builder.addIncludeCell("A6");
		builder.addIncludeCell("B6");
		builder.addIncludeCell("C6");
		builder.addIncludeCell("A7");
		builder.addIncludeCell("B7");
		builder.addIncludeCell("C7");
		
		ExcelToCanvas excel1 = builder.build(f);
		List<StrInfo> list1 = excel1.getStrs();
		assertTrue(list1.size() >= 12);
		
		for (StrInfo info : list1) {
			assertFalse("Exists B column", info.getId().indexOf("B") != -1);
		}
		builder.setIncludeHiddenCell(true);
		ExcelToCanvas excel2 = builder.build(f);
		List<StrInfo> list2 = excel2.getStrs();
		assertTrue(list2.size() >= 18);
		
		int cnt = 0;
		for (StrInfo info : list2) {
			if (info.getId().indexOf("B") == 0) {
				cnt++;
			}
		}
		assertTrue("Invalid B column count", cnt == 6);
	}
}
