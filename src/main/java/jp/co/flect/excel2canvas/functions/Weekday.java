package jp.co.flect.excel2canvas.functions;


import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;

import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.formula.eval.EvaluationException;
import org.apache.poi.ss.formula.eval.NumberEval;
import org.apache.poi.ss.formula.eval.OperandResolver;
import org.apache.poi.ss.formula.eval.ValueEval;
import org.apache.poi.ss.formula.functions.Function1Arg;
import org.apache.poi.ss.formula.functions.Function2Arg;
import org.apache.poi.ss.usermodel.DateUtil;

/**
 * Function - WEEKDAY<br>
 * The logic of this class is copied from org.apache.poi.ss.formula.functions.CalendarFieldFunction
 */
public class Weekday implements Function1Arg, Function2Arg {
	
	private final int _dateFieldId;
	
	public Weekday() {
		this._dateFieldId = Calendar.DAY_OF_WEEK;
	}
	
	public final ValueEval evaluate(ValueEval[] args, int srcRowIndex, int srcColumnIndex) {
		switch (args.length) {
			case 1:
				return evaluate(srcRowIndex, srcColumnIndex, args[0], null);
			case 2:
				return evaluate(srcRowIndex, srcColumnIndex, args[0], args[1]);
		}
		return ErrorEval.VALUE_INVALID;
	}
	
	public final ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0) {
		return evaluate(srcRowIndex, srcColumnIndex, arg0, null);
	}
	
	public final ValueEval evaluate(int srcRowIndex, int srcColumnIndex, ValueEval arg0, ValueEval arg1) {
		double val;
		int type = 1;
		try {
			ValueEval ve = OperandResolver.getSingleValue(arg0, srcRowIndex, srcColumnIndex);
			val = OperandResolver.coerceValueToDouble(ve);
			if (arg1 != null) {
				ValueEval ve2 = OperandResolver.getSingleValue(arg1, srcRowIndex, srcColumnIndex);
				type = (int)OperandResolver.coerceValueToDouble(ve2);
			}
		} catch (EvaluationException e) {
			return e.getErrorEval();
		}
		if (val < 0) {
			return ErrorEval.NUM_ERROR;
		}
		int ret = getCalField(val);
		if (type != 1) {
			switch (type) {
				case 2:
					ret = ret == 1 ? 7 : ret - 1;
					break;
				case 3:
					ret--;
					break;
				default:
					return ErrorEval.NUM_ERROR;
			}
		}
		return new NumberEval(ret);
	}
	
	private int getCalField(double serialDate) {
		// For some reason, a date of 0 in Excel gets shown
		//  as the non existant 1900-01-00
		if(((int)serialDate) == 0) {
			switch (_dateFieldId) {
				case Calendar.YEAR: return 1900;
				case Calendar.MONTH: return 1;
				case Calendar.DAY_OF_MONTH: return 0;
			}
			// They want time, that's normal
		}
		
		// TODO Figure out if we're in 1900 or 1904
		Date d = DateUtil.getJavaDate(serialDate, false);
		
		Calendar c = new GregorianCalendar();
		c.setTime(d);
		int result = c.get(_dateFieldId);
		
		// Month is a special case due to C semantics
		if (_dateFieldId == Calendar.MONTH) {
			result++;
		}
		
		return result;
	}
}
