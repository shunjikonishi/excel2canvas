package jp.co.flect.excel2canvas.validator;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTExtensionList;

import java.util.Collections;
import java.util.List;
import java.util.ArrayList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;

public class ExtDataValidationExtractor {

	public List<Element> getDataValidationNode(XSSFSheet sheet) {
		CTWorksheet ctWorksheet = sheet.getCTWorksheet();
		CTExtensionList extList = ctWorksheet.getExtLst();
		if (extList == null) {
			return Collections.<Element>emptyList();
		}
		List<Element> list = new ArrayList<Element>();
		Node node = extList.getDomNode().getFirstChild();
		while (node != null) {
			if ("ext".equals(node.getLocalName())) {
				processExt((Element)node, list);
			}
			node = node.getNextSibling();
		}
		return list;
	}

	private void processExt(Element el, List<Element> list) {
		Node node = el.getFirstChild();
		while (node != null) {
			if ("dataValidations".equals(node.getLocalName())) {
				processDataValidations((Element)node, list);
			}
			node = node.getNextSibling();
		}
	}

	private void processDataValidations(Element el, List<Element> list) {
		Node node = el.getFirstChild();
		while (node != null) {
			if ("dataValidation".equals(node.getLocalName())) {
				Element dv = (Element)node;
				if ("list".equals(dv.getAttribute("type"))) {
					list.add(dv);
				}
			}
			node = node.getNextSibling();
		}
	}
}