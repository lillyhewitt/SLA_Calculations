package main;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

// different styles used across multiple classes, do not change, just add styles and change which style is called in the separate class
public class ExcelStyleUtil {

	public static byte[] createBlue() {
		byte[] rgbBlue = new byte[]{19, 43, 81};
		return rgbBlue;
	}

	public static byte[] createGray() {
		byte[] rgbGray = new byte[]{(byte) 211, (byte) 211, (byte) 211};
		return rgbGray;
	}

	public static byte[] createYellow() {
		byte[] rgbYellow = new byte[]{(byte) 255, (byte) 255, 0};
		return rgbYellow;
	}

	public static XSSFCellStyle createNavyBlueStyle(XSSFWorkbook workbook) {
		XSSFCellStyle navyBlueStyle = workbook.createCellStyle();
		XSSFColor navyBlue = new XSSFColor(createBlue(), new DefaultIndexedColorMap());
		navyBlueStyle.setFillForegroundColor(navyBlue);
		navyBlueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		navyBlueStyle.setAlignment(HorizontalAlignment.CENTER);
		navyBlueStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		navyBlueStyle.setWrapText(true); 

		XSSFFont font = workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex()); // Change font color to white
		font.setFontHeightInPoints((short) 18); // Change font size here
		font.setBold(true);
		navyBlueStyle.setFont(font);

		return navyBlueStyle;
	}

	public static XSSFCellStyle createSLAheaders(XSSFWorkbook workbook) {
		XSSFCellStyle slaHeaders = workbook.createCellStyle();
		slaHeaders.setAlignment(HorizontalAlignment.CENTER);
		slaHeaders.setVerticalAlignment(VerticalAlignment.CENTER);
		slaHeaders.setWrapText(true);

		return slaHeaders;
	}

	public static XSSFCellStyle createSLAdescriptors(XSSFWorkbook workbook) {
		XSSFCellStyle slaHeaders = workbook.createCellStyle();
		slaHeaders.setWrapText(true);
		slaHeaders.setBorderTop(BorderStyle.DOTTED);
		slaHeaders.setTopBorderColor(IndexedColors.BLACK.getIndex());
		slaHeaders.setAlignment(HorizontalAlignment.LEFT);
		slaHeaders.setVerticalAlignment(VerticalAlignment.CENTER);

		return slaHeaders;
	}


	public static XSSFCellStyle sideSLAdescriptors(XSSFWorkbook workbook) {
		XSSFCellStyle slaHeaders = workbook.createCellStyle();
		slaHeaders.setWrapText(true);
		slaHeaders.setBorderTop(BorderStyle.DOTTED);
		slaHeaders.setTopBorderColor(IndexedColors.BLACK.getIndex());
		slaHeaders.setAlignment(HorizontalAlignment.CENTER);
		slaHeaders.setVerticalAlignment(VerticalAlignment.CENTER);
		slaHeaders.setWrapText(true);

		return slaHeaders;
	}

	public static XSSFCellStyle createGrayHeaderStyle(XSSFWorkbook workbook) {
		XSSFCellStyle grayHeaderStyle = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		grayHeaderStyle.setFillForegroundColor(borderColor);
		grayHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		grayHeaderStyle.setWrapText(true);
		grayHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
		grayHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		grayHeaderStyle.setBorderBottom(BorderStyle.MEDIUM);
		grayHeaderStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
		grayHeaderStyle.setBorderTop(BorderStyle.MEDIUM);
		grayHeaderStyle.setTopBorderColor(IndexedColors.RED.getIndex());
		grayHeaderStyle.setBorderLeft(BorderStyle.MEDIUM);
		grayHeaderStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
		grayHeaderStyle.setBorderRight(BorderStyle.MEDIUM);
		grayHeaderStyle.setRightBorderColor(IndexedColors.RED.getIndex());
		grayHeaderStyle.setVerticalAlignment(VerticalAlignment.TOP);

		return grayHeaderStyle;
	}

	public static XSSFCellStyle createPlainTableHeaderStyle(XSSFWorkbook workbook) {
		XSSFCellStyle plaintableHeaderStyle = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		plaintableHeaderStyle.setFillForegroundColor(borderColor);
		plaintableHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		plaintableHeaderStyle.setBorderBottom(BorderStyle.MEDIUM);
		plaintableHeaderStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		plaintableHeaderStyle.setBorderLeft(BorderStyle.MEDIUM);
		plaintableHeaderStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		plaintableHeaderStyle.setBorderRight(BorderStyle.MEDIUM);
		plaintableHeaderStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		plaintableHeaderStyle.setBorderTop(BorderStyle.MEDIUM);
		plaintableHeaderStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
		plaintableHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
		plaintableHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);

		return plaintableHeaderStyle;
	}

	public static XSSFCellStyle createResumeHeader(XSSFWorkbook workbook) {
		XSSFCellStyle resumeHeader = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		resumeHeader.setFillForegroundColor(borderColor);
		resumeHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		resumeHeader.setWrapText(true);
		resumeHeader.setAlignment(HorizontalAlignment.CENTER);
		resumeHeader.setVerticalAlignment(VerticalAlignment.CENTER);
		resumeHeader.setBorderBottom(BorderStyle.MEDIUM);
		resumeHeader.setBottomBorderColor(IndexedColors.BLACK.getIndex());

		return resumeHeader;
	}

	public static XSSFCellStyle createResumeGray(XSSFWorkbook workbook) {
		XSSFCellStyle resumeGray = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		resumeGray.setFillForegroundColor(borderColor);
		resumeGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		resumeGray.setWrapText(true);
		resumeGray.setVerticalAlignment(VerticalAlignment.TOP);

		return resumeGray;
	}

	public static XSSFCellStyle createHeaderStyle(XSSFWorkbook workbook) {
		XSSFCellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		headerStyle.setAlignment(HorizontalAlignment.CENTER); 
		headerStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		headerStyle.setWrapText(true); 

		XSSFFont fontHeader = workbook.createFont();
		fontHeader.setColor(IndexedColors.DARK_TEAL.getIndex());
		fontHeader.setFontHeightInPoints((short) 12); // Change font size here
		fontHeader.setBold(true);
		headerStyle.setFont(fontHeader);
		headerStyle.setWrapText(true);

		return headerStyle;
	}

	public static XSSFCellStyle createFraudHeaderStyle(XSSFWorkbook workbook) {
		// Create borders style
		XSSFCellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setWrapText(true);
		XSSFFont fontHeader = workbook.createFont();
		fontHeader.setColor(IndexedColors.DARK_TEAL.getIndex());
		fontHeader.setFontHeightInPoints((short) 12); // Change font size here
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		headerStyle.setFont(fontHeader);

		return headerStyle;
	}

	public static XSSFCellStyle createSideHeaderStyle(XSSFWorkbook workbook) {
		XSSFCellStyle sideHeaderStyle = workbook.createCellStyle();
		XSSFColor sideHeaderColor = new XSSFColor(createBlue(), new DefaultIndexedColorMap());
		sideHeaderStyle.setFillForegroundColor(sideHeaderColor);
		sideHeaderStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		sideHeaderStyle.setAlignment(HorizontalAlignment.CENTER); 
		sideHeaderStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER); 
		sideHeaderStyle.setWrapText(true);

		return sideHeaderStyle;
	}

	public static XSSFCellStyle createBorderGrayStyle(XSSFWorkbook workbook) {
		XSSFCellStyle borderGrayStyle = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		borderGrayStyle.setFillForegroundColor(borderColor);
		borderGrayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		borderGrayStyle.setAlignment(HorizontalAlignment.CENTER);
		borderGrayStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		XSSFFont sideFont = workbook.createFont();
		sideFont.setFontHeightInPoints((short) 20); // Change font size here
		borderGrayStyle.setFont(sideFont);

		return borderGrayStyle;
	}

	public static XSSFCellStyle createPlainTableStyle(XSSFWorkbook workbook) {
		XSSFCellStyle tableStyle = workbook.createCellStyle();
		tableStyle.setBorderBottom(BorderStyle.THIN);
		tableStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		tableStyle.setBorderLeft(BorderStyle.THIN);
		tableStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		tableStyle.setBorderRight(BorderStyle.THIN);
		tableStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		tableStyle.setBorderTop(BorderStyle.THIN);
		tableStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

		return tableStyle;
	}

	public static XSSFCellStyle createGrayStyle(XSSFWorkbook workbook) {
		XSSFCellStyle grayStyle = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		grayStyle.setFillForegroundColor(borderColor);
		grayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		grayStyle.setWrapText(true);
		grayStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		grayStyle.setBorderBottom(BorderStyle.MEDIUM);
		grayStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setBorderTop(BorderStyle.MEDIUM);
		grayStyle.setTopBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setBorderLeft(BorderStyle.MEDIUM);
		grayStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setBorderRight(BorderStyle.MEDIUM);
		grayStyle.setRightBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setVerticalAlignment(VerticalAlignment.TOP);

		return grayStyle;
	}

	public static XSSFCellStyle createGrayResumeFraudStyle(XSSFWorkbook workbook) {
		XSSFCellStyle grayStyle = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		grayStyle.setFillForegroundColor(borderColor);
		grayStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		grayStyle.setWrapText(true);
		grayStyle.setAlignment(HorizontalAlignment.CENTER);
		grayStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		grayStyle.setBorderBottom(BorderStyle.MEDIUM);
		grayStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setBorderTop(BorderStyle.MEDIUM);
		grayStyle.setTopBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setBorderLeft(BorderStyle.MEDIUM);
		grayStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
		grayStyle.setBorderRight(BorderStyle.MEDIUM);
		grayStyle.setRightBorderColor(IndexedColors.RED.getIndex());

		return grayStyle;
	}

	public static XSSFCellStyle createNavyTableStyle(XSSFWorkbook workbook) {
		XSSFCellStyle navyTableStyle = workbook.createCellStyle();
		XSSFColor navyBlue = new XSSFColor(createBlue(), new DefaultIndexedColorMap());
		navyTableStyle.setFillForegroundColor(navyBlue);
		navyTableStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		navyTableStyle.setAlignment(HorizontalAlignment.CENTER);
		navyTableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		navyTableStyle.setWrapText(true); // Enable text wrapping

		XSSFFont fontTable = workbook.createFont();
		fontTable.setColor(IndexedColors.WHITE.getIndex()); // Change font color to white
		fontTable.setFontHeightInPoints((short) 12); // Change font size here
		fontTable.setBold(true);
		navyTableStyle.setFont(fontTable);

		return navyTableStyle;
	}

	public static XSSFCellStyle createTableStyle(XSSFWorkbook workbook) {
		XSSFCellStyle tableStyle = workbook.createCellStyle();
		XSSFColor borderColor = new XSSFColor(createGray(), new DefaultIndexedColorMap());
		tableStyle.setFillForegroundColor(borderColor);
		tableStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		tableStyle.setAlignment(HorizontalAlignment.LEFT);
		tableStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		tableStyle.setBorderBottom(BorderStyle.THIN);
		tableStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		tableStyle.setBorderLeft(BorderStyle.THIN);
		tableStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		tableStyle.setBorderRight(BorderStyle.THIN);
		tableStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
		tableStyle.setBorderTop(BorderStyle.THIN);
		tableStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

		return tableStyle;
	}

	public static XSSFCellStyle createYellowStyle(XSSFWorkbook workbook) {
		XSSFCellStyle yellowStyle = workbook.createCellStyle();
		XSSFColor yellow = new XSSFColor(createYellow(), new DefaultIndexedColorMap());
		yellowStyle.setFillForegroundColor(yellow);
		yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		yellowStyle.setWrapText(true);
		yellowStyle.setBorderBottom(BorderStyle.MEDIUM);
		yellowStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
		yellowStyle.setBorderTop(BorderStyle.MEDIUM);
		yellowStyle.setTopBorderColor(IndexedColors.RED.getIndex());
		yellowStyle.setBorderLeft(BorderStyle.MEDIUM);
		yellowStyle.setLeftBorderColor(IndexedColors.RED.getIndex());
		yellowStyle.setBorderRight(BorderStyle.MEDIUM);
		yellowStyle.setRightBorderColor(IndexedColors.RED.getIndex());
		yellowStyle.setAlignment(HorizontalAlignment.CENTER);
		yellowStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

		return yellowStyle;
	}

	public static XSSFCellStyle createBrownStyle(XSSFWorkbook workbook) {
		XSSFCellStyle brownStyle = workbook.createCellStyle();
		byte[] rbgBrown = new byte[]{(byte) 216, (byte) 207, (byte) 196};
		XSSFColor brown = new XSSFColor(rbgBrown, new DefaultIndexedColorMap());
		brownStyle.setFillForegroundColor(brown); 
		brownStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		return brownStyle;
	}

	public static XSSFCellStyle redBackground(XSSFWorkbook workbook) {
		XSSFCellStyle redStyle = workbook.createCellStyle();
		XSSFColor red = new XSSFColor(new byte[]{(byte) 255, 0, 0}, new DefaultIndexedColorMap());
		redStyle.setFillForegroundColor(red); 
		redStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		redStyle.setWrapText(true); // Enable text wrapping
		redStyle.setAlignment(HorizontalAlignment.CENTER);
		redStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

		return redStyle;
	}

	public static XSSFCellStyle greenBackground(XSSFWorkbook workbook) {
		XSSFCellStyle greenStyle = workbook.createCellStyle();
		XSSFColor green = new XSSFColor(new byte[]{ 0, (byte) 150, 0}, new DefaultIndexedColorMap());
		greenStyle.setFillForegroundColor(green); 
		greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		greenStyle.setWrapText(true); // Enable text wrapping
		greenStyle.setAlignment(HorizontalAlignment.CENTER);
		greenStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

		return greenStyle;
	}

	public static XSSFCellStyle yellowBackground(XSSFWorkbook workbook) {
		XSSFCellStyle yellowStyle = workbook.createCellStyle();
		XSSFColor yellow = new XSSFColor(new byte[]{ (byte) 255, (byte) 255, 0}, new DefaultIndexedColorMap());
		yellowStyle.setFillForegroundColor(yellow); 
		yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		yellowStyle.setWrapText(true); // Enable text wrapping
		yellowStyle.setAlignment(HorizontalAlignment.CENTER);
		yellowStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

		return yellowStyle;
	}

	public static XSSFCellStyle createBlueStyle(XSSFWorkbook workbook) {
		XSSFCellStyle blueStyle = workbook.createCellStyle();
		byte[] rbgBlue = new byte[]{(byte) 175, (byte) 227, (byte) 233};
		XSSFColor blue = new XSSFColor(rbgBlue, new DefaultIndexedColorMap());
		blueStyle.setFillForegroundColor(blue); 
		blueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		blueStyle.setAlignment(HorizontalAlignment.CENTER);
		blueStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
		blueStyle.setWrapText(true); 

		XSSFFont fontTable = workbook.createFont();
		fontTable.setFontHeightInPoints((short) 12); // Change font size here
		fontTable.setBold(true);
		blueStyle.setFont(fontTable);

		return blueStyle;
	}

	public static XSSFFont createUnderlineStyle(XSSFWorkbook workbook) {
		XSSFFont underlineFont = workbook.createFont();
		underlineFont.setUnderline(Font.U_SINGLE);
		underlineFont.setBold(true);
		underlineFont.setColor(IndexedColors.RED.getIndex());

		return underlineFont;
	}

	public static XSSFFont createBoldStyle(XSSFWorkbook workbook) {
		XSSFFont boldFont = workbook.createFont();
		boldFont.setBold(true);

		return boldFont;
	}

	public static XSSFFont createRedStyle(XSSFWorkbook workbook) {
		XSSFFont redFont = workbook.createFont();
		redFont.setColor(IndexedColors.RED.getIndex());

		return redFont;
	}

}

