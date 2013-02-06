package jp.co.flect.excel2canvas.functions;

import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.functions.Function;

/**
 * Register costom function to POI.
 */
public class FunctionManager {
	
	private static boolean registered = false;
	
	public static void registerAll() {
		synchronized (FunctionManager.class) {
			if (!registered) {
				register("WEEKDAY", new Weekday());
				registered = true;
			}
		}
	}
	
	private static void register(String name, Function func) {
		try {
			FunctionEval.registerFunction("WEEKDAY", new jp.co.flect.excel2canvas.functions.Weekday());
		} catch (IllegalArgumentException e) {
			//Ignore
		}
	}
}
