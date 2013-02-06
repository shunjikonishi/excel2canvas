package jp.co.flect.excel2canvas;

import java.awt.image.BufferedImage;

import jp.co.flect.excel2canvas.ExcelToCanvas.FillInfo;
import jp.co.flect.excel2canvas.ExcelToCanvas.LineInfo;
import jp.co.flect.excel2canvas.ExcelToCanvas.PictureInfo;
import jp.co.flect.excel2canvas.ExcelToCanvas.StrInfo;

/**
 * !!! Not implemented yet.<br>
 * Convert ExcelToCanvas to the image.
 */
public class ImageWriter {
	
	public BufferedImage write(ExcelToCanvas excel) {
		BufferedImage img = new BufferedImage(excel.getWidth(), excel.getHeight(), BufferedImage.TYPE_INT_ARGB);
		writeFills(img, excel);
		//writeLines(img, excel);
		//writePictures(img, excel);
		//writeStrs(img, excel);
		return img;
	}
	
	private void writeFills(BufferedImage img, ExcelToCanvas excel) {
		for (FillInfo info : excel.getFills()) {
			writeFills(img, info);
		}
	}
	
	private void writeFills(BufferedImage img, FillInfo info) {
		//ToDo
	}
	
}
