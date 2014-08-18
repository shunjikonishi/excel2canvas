package jp.co.flect.excel2canvas;

import java.util.Collections;
import java.util.List;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellReference;

/**
 * 名前つきセルの情報
 * CellListには結合セルを除いた実セルのみを含む
 */
public class NamedCellInfo {
	
	private Name name;
	private List<CellReference> list;
	
	public NamedCellInfo(Name name, List<CellReference> list) {
		this.name = name;
		this.list = list;
	}

	public Name getPoiName() { return this.name;}
	
	public String getName() { return this.name.getNameName();}
	public String getRange() { return this.name.getRefersToFormula();}
	public String getSheetName() { return this.name.getSheetName();}
	public String getComment() { return this.name.getComment();}
	
	public int getCellCount() { return this.list.size();}
	public List<CellReference> getCellList() { return this.list;}
	
	public List<CellReference> getCellListWithOffset(int offset) {
		return offset == 0 ? this.list : offset >= this.list.size() ? Collections.<CellReference>emptyList() : this.list.subList(offset, this.list.size());
	}
}
