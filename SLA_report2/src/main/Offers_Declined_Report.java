package main;

import org.apache.poi.ss.usermodel.Row;
import java.io.File;
import java.io.FileInputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Offers_Declined_Report {
	// create variables to store KPI, declined offers, and num of offers in the quarter
	public static String KPI = "";
	public static int declinedOffers = 0;
	public static int numOffers = 0;

	// Read from an Excel file
	public static void read_From_Excel(String fileName, int quarter, int year) {
		try {
			// Creating a new file instance
			File file = new File(fileName);
			// Obtaining bytes from the file
			FileInputStream fis = new FileInputStream(file);
			// Creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);

			// loop through cells in sheet 1 (Open) and sheet 2 (Closed)
			for (int i = 1; i < 3; i++) {
				XSSFSheet sheet = wb.getSheetAt(i);

				// index through rows in the sheet
				for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
					Row row = sheet.getRow(rowIndex);
					if (row == null) {
						break; // Skip if the row is null
					}

					// only process rows in active status req table in open
					Cell dateCell = row.getCell(9);  // Column J
					if(i == 1) {
						if (dateCell.getCellType() != CellType.NUMERIC) {
							rowIndex = sheet.getLastRowNum();
						} else {
							processRow(row, quarter, year, false);
						}
					}else if (dateCell != null && dateCell.getDateCellValue() != null) {
						processRow(row, quarter, year, true);
					}
				}
			}

			// add number of declined offers to number of offers
			numOffers += declinedOffers;
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void processRow(Row row, int quarter, int year, boolean isClosedSheet) {
		Cell dateCell = isClosedSheet ? row.getCell(34) : row.getCell(9);  // Column AI for closed sheet, Column J for open sheet
		Cell vendorCell = row.getCell(13);  // Column N
		Cell statusCell = isClosedSheet ? row.getCell(37) : null;  // Column AL
		Cell declineReasonCell = row.getCell(24);  // Column Y
		Cell offerCell = isClosedSheet ? row.getCell(38) : null;  // Column AM for closed sheet

		Date date = QuarterCheck.getDateFromCell(dateCell);
		if (date != null && QuarterCheck.isInCorrectQuarter(date, quarter, year)) {
			if (VendorCheck.isValidVendor(vendorCell)) {
				// increase number of declined offers if status is not cancelled and declined reason is found in Close tab
				if (isClosedSheet && statusCell != null && !"C".equals(statusCell.getStringCellValue()) && containsDecline(declineReasonCell)) {
					declinedOffers++;
				}
				// increase number of declined offers if declined reason is found in Open tab
				else if(!isClosedSheet && containsDecline(declineReasonCell)) {
					declinedOffers++;
				}
				// increase number of offers if status is not cancelled and offer is filled by Entech in Close tab
				if (offerCell != null && statusCell != null && !"C".equals(statusCell.getStringCellValue()) && "E".equals(offerCell.getStringCellValue())) {
					numOffers++;
				}
			}
		}
	}

	// Check if cell contains the word "decline"
	private static boolean containsDecline(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String cellValue = cell.getStringCellValue().toLowerCase();
			return cellValue.contains("decline");
		}
		return false;
	}

	public static String calculateKPI(int quarter, int year, String filePath) {
		KPI = "";
		declinedOffers = 0;
		numOffers = 0;
		
		// Get declined offers and total offers from File 
		read_From_Excel(filePath, quarter, year);
		// Calculate KPI
		double ratio = ((double) declinedOffers/numOffers) * 100;
		String ratioFormatted = String.format("%.2f", ratio) + "%";
		KPI = ratioFormatted;
		return KPI;
	}
	
	// Write Ratio Resumes to Interviews Exc to Excel file
	public void write_Offers_Declined(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year, String filePath) {
		// Add merged regions for blue KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 4));
		sheet.addMergedRegion(new CellRangeAddress(64, 65, 0, 4));
		sheet.addMergedRegion(new CellRangeAddress(2, 63, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 63, 4, 4));

		// set title for KPI box and style
		for (int i = 0; i < 65; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 0; j < 5; j++) { 
				Cell cell = row.getCell(j);
				if (cell == null) {
					cell = row.createCell(j);
					if(i == 0 && j == 0) {
						cell.setCellValue("KPI Calculation");
					}
				}
				cell.setCellStyle(ExcelStyleUtil.createNavyBlueStyle(workbook));
			}
		}

		// Add merged regions for gray KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(2, 4, 1, 3));
		sheet.addMergedRegion(new CellRangeAddress(5, 63, 1, 3));

		// Create text for gray KPI Calc box
		for (int i = 2; i < 63; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 1; j < 4; j++) {
				Cell cell = row.getCell(1);
				if (cell == null) {
					cell = row.createCell(j);
				}
				if (i == 2) {
					cell.setCellValue("Number of offers declined \n(exclusive and non-exclusive)\n**excluding cancelled reqs]"); 
				}
				else if (i == 5) { // Correct the condition to apply to the correct row
					// Create a rich text string for the fifth row
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE # DECLINED\nFilter from the \"Closed\" tab:\n" +
									"1) Column N (Vendor) ->\n filter by \"All\" and vendor name \n(ex: All + JTS)\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"3) Column AL (Filled/Cancelled) ->\n unselect \'C\' (only filled reqs remain)\n" +
									"4) Based on filtered data, determine if\n there are any offers for vendor\n candidates that were declined\n" +
									"**Best practice: highlight column Y\n(Status) -> ctrl+F -> search \"decline\"\n-> identity declined offers per vendor\n" +
									"5) Enter # of declined offers identified\ninto KPI Calculation" +
									"\n\nFilter from the \"Open\" tab:\n" +
									"1) Column J (Opened) ->\n filter by SLA quarter\n(ex: All + JTS)\n" +
									"2) Column N (Vendor) ->\n filter by \"All\" and vendor name \n(ex: All + JTS\n" +
									"3) Based on filtered data, determine if\n there are any offers for vendor\n candidates that were declined\n" +
									"**Best practice: highlight column Y\n(Status) -> ctrl+F -> search \"decline\"\n-> identity declined offers per vendor\n" +
									"4) Add # of declined offers identified\nfrom the \"Open\" tab to the # of\n declined offers identified\nfrom the \"Closed\" tab in the KPI Calculation"+
									"\n\nTO DETERMINE # OFFERS\nFilter from the \"Closed\" tab:\n" +
									"1) Column N (Vendor) ->\n filter by \"All\" and vendor name \n(ex: All + JTS)\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"3) Column AL (Filled/Cancelled) ->\n unselect \'C\' (only filled reqs remain)\n" +
									"4) Column AM (Filled by) ->\nleave only your vendor selected\n" +
									"5) Enter # of positions closed + number of\noffers declined into the KPI Calculation\n"+
									"**(ex: per filtered data, if RT filled 55\nreqs in Q3 AND had 1 declined offer =\n 56 total offers)"
							);
					// Apply the bold style to the specified characters
					richString.applyFont(0, 25, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(25, 55, ExcelStyleUtil.createBoldStyle(workbook));
					richString.applyFont(597, 626, ExcelStyleUtil.createBoldStyle(workbook));
					richString.applyFont(1125, 1146, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(1147, 1175, ExcelStyleUtil.createBoldStyle(workbook));
					// Set the cell value to the rich text string
					cell.setCellValue(richString);
				}
				sheet.setColumnWidth(j, 256 * 20);
			}
		}

		// Set style for gray KPI Calc box
		CellRangeAddress mergedRegion = new CellRangeAddress(2, 63, 1, 3);
		for (int rowNum = mergedRegion.getFirstRow(); rowNum <= mergedRegion.getLastRow(); rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				row = sheet.createRow(rowNum);
			}
			for (int colNum = mergedRegion.getFirstColumn(); colNum <= mergedRegion.getLastColumn(); colNum++) {
				Cell cell = row.getCell(colNum);
				if (cell == null) {
					cell = row.createCell(colNum);
				}
				else if(rowNum == 2) {
					cell.setCellStyle(ExcelStyleUtil.createGrayHeaderStyle(workbook));
				}
				else {
					cell.setCellStyle(ExcelStyleUtil.createGrayStyle(workbook));
				}
			}
		}

		// Add merged regions ratio of resumes to interviews table
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 6, 7));

		// Add text and style for ratio of resumes to interviews table
		for(int r = 0; r < 5; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				row = sheet.createRow(r);
			}
			Cell cell = row.createCell(6);
			if(r == 0) {
				cell.setCellValue("Number of Offers Declined\nKPI Calculation");
				cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
			}
			else if(r == 2) {
				cell.setCellValue("# declined");
				cell.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(declinedOffers);
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
			else if(r == 3) {
				cell.setCellValue("# offers");
				cell.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(numOffers);
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
			else if(r == 4) {
				cell.setCellValue("Ratio:");
				cell.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(calculateKPI(quarter, year, filePath));
				cell2.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
			}
		}
		sheet.setColumnWidth(6, 256 * 10);
		sheet.setColumnWidth(7, 256 * 10);

	}
}
