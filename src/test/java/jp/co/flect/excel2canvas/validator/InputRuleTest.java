package jp.co.flect.excel2canvas.validator;

import static org.apache.poi.ss.usermodel.DataValidationConstraint.OperatorType;
import static org.apache.poi.ss.usermodel.DataValidationConstraint.ValidationType;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.fail;
import org.junit.Test;

import jp.co.flect.excel2canvas.ExcelUtils;

public class InputRuleTest {

	@Test
	public void validateTest() throws Exception {
		InputRule rule1 = InputRule.forTest(ValidationType.DATE, OperatorType.GREATER_THAN, "2014/1/1", null);
		rule1.validate("2014-08-01");
		try {
			rule1.validate("2013-12-31");
			fail();
		} catch (Exception e) {
		}
		InputRule rule2 = InputRule.forTest(ValidationType.DATE, OperatorType.GREATER_THAN, "41640", null);
		rule2.validate("2014-08-01");
		try {
			rule2.validate("2013-12-31");
			fail();
		} catch (Exception e) {
		}
	}

	public void numericTest() {
		String[] ok_strs = {
			"#,##0",
			"0",
			"@",
			"\"$\"#,##0.00_);(\"$\"#,##0.00)"
		};
		for (String s : ok_strs) {
			assertTrue(s, ExcelUtils.isNumericStyle(s));
		}
	}
}