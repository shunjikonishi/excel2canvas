package jp.co.flect.excel2canvas;

import java.io.File;
import java.util.List;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.fail;
import org.junit.Test;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;

import jp.co.flect.excel2canvas.ExcelToCanvas.StrInfo;

/**
 * Unit test for simple App.
 */
public class InvoiceTest {
	
	@Test
	/**
	 * B列が非表示列のシートを読み込むテスト
	 */
	public void compareFromTo() throws Exception {
		File f = new File("testdata/Invoice.xlsx");
		Workbook workbook = ExcelUtils.createWorkbook(f);
		List<NamedCellInfo> list = ExcelUtils.createNamedCellList(workbook);
		assertEquals(12, list.size());
		for (NamedCellInfo name: list) {
			System.out.println("name: " + name.getName() + ", " + name.getSheetName());
			Sheet sheet = workbook.getSheet(name.getSheetName());
			assertNotNull(name.getSheetName(), sheet);
		}
	}
}
