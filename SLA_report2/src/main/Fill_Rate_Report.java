package main;

import org.apache.poi.ss.usermodel.Row;
import java.io.File;
import java.io.FileInputStream;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Fill_Rate_Report {
	// CHANGE to jobs received only in this quarter - maybe

	// create variables to store KPI calculations
	public static String KPIexc = "";
	public static String KPInon = "";
	public static int jobsFilledExc = 0;
	public static int jobsReceivedExc = 0;
	public static int jobsFilledNon = 0;
	public static int jobsReceivedNon = 0;

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

					// only check entries in active status req table for open sheet
					Cell dateCell = row.getCell(9);  // Column J
					if(i == 1) {
						if (dateCell.getCellType() != CellType.NUMERIC) {
							rowIndex = sheet.getLastRowNum();
						} else {
							processRow(row, quarter, year, false);
						}
					} else if (dateCell != null && dateCell.getDateCellValue() != null) {
						processRow(row, quarter, year, true);
					}
				}
			}

			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Find jobs filled and received this quarter by checking dates, vendor, and offer cells of each entry in the roster
	private static void processRow(Row row, int quarter, int year, boolean isClosedSheet) {
		Cell closeCell = isClosedSheet ? row.getCell(34) : null;  // Column AI 
		Cell openedCell = row.getCell(9); // Column J
		Cell vendorCell = row.getCell(13);  // Column N
		Cell offerCell = isClosedSheet ? row.getCell(38) : null; // Column AM 

		// check close date for this quarter
		Date date = isClosedSheet ? QuarterCheck.getDateFromCell(closeCell) : null;
		if (date != null && QuarterCheck.isInCorrectQuarter(date, quarter, year)) {
			// check if vendor is Etnech
			if( VendorCheck.isJustVendor(vendorCell)) {
				// increase exclusive job filled when offer cell is "E" and just Entech with an end date in this quarter
				if(offerCell != null && "E".equals(offerCell.getStringCellValue())) {
					jobsFilledExc++;
				}
				Date openDate = QuarterCheck.getDateFromCell(openedCell);
				// increase exclusive job received when just Entech with an opened date in this quarter
				if (openDate != null && QuarterCheck.isInCorrectQuarter(openDate, quarter, year)) {
					jobsReceivedExc++;
				}
			}
			// check if vendor is All
			if( VendorCheck.isAllVendor(vendorCell)) {
				// increase nonexclusive job filled when offer cell is "E" and All vendors with an end date in this quarter
				if(offerCell != null && "E".equals(offerCell.getStringCellValue())) {
					jobsFilledNon++;
				}
				Date openDate = QuarterCheck.getDateFromCell(openedCell);
				// increase exclusive job received when All vendors with an opened date in this quarter
				if (openDate != null && QuarterCheck.isInCorrectQuarter(openDate, quarter, year)) {
					jobsReceivedNon++;
				}
			}
		}
		if(!isClosedSheet) { // row in an open sheet
			Date openDate = QuarterCheck.getDateFromCell(openedCell);
			// check if open date is in this quarter
			if (openDate != null && QuarterCheck.isInCorrectQuarter(openDate, quarter, year)) {
				// increase exclusive job filled when just Entech with an open date in this quarter
				if(VendorCheck.isJustVendor(vendorCell) ) {
					jobsReceivedExc++;
				}
				// increase nonexclusive job filled when All vendors with an open date in this quarter
				if (VendorCheck.isAllVendor(vendorCell)) {
					jobsReceivedNon++;
				}
			}
		}
	}

	public static String calculateExcKPI(int quarter, int year, String filePath) {
		// reset variables
		KPIexc = "";
		jobsFilledExc = 0;
		jobsReceivedExc = 0;
		KPInon = "";
		jobsFilledNon = 0;
		jobsReceivedNon = 0;

		// Get jobs filled and received
		read_From_Excel(filePath, quarter, year);
		// Calculate KPI
		double ratioExc = ((double) jobsFilledExc/jobsReceivedExc) * 100;
		String ratioExcFormatted = String.format("%.2f", ratioExc) + "%";
		KPIexc = ratioExcFormatted;
		return KPIexc;
	}

	public static String calculateNonKPI(int quarter, int year, String filePath) {
		// Calculate KPI
		double ratioNon = ((double) jobsFilledNon/jobsReceivedNon) * 100;
		String ratioNonFormatted = String.format("%.2f", ratioNon) + "%";
		KPInon = ratioNonFormatted;
		return KPInon;
	}

	// Write Ratio Resumes to Interviews Exc to Excel file
	public void write_Fill_Rate(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year, String filePath) {
		// Add merged regions for blue KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 3));
		sheet.addMergedRegion(new CellRangeAddress(38, 39, 0, 3));
		sheet.addMergedRegion(new CellRangeAddress(2, 37, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 37, 3, 3));
		sheet.addMergedRegion(new CellRangeAddress(40, 76, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(40, 76, 3, 3));
		sheet.addMergedRegion(new CellRangeAddress(77, 78, 0, 3));

		// set title for KPI box and style
		for (int i = 0; i < 78; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 0; j < 4; j++) { 
				Cell cell = row.getCell(j);
				if (cell == null) {
					cell = row.createCell(j);
					if(i == 0 && j == 0) {
						cell.setCellValue("KPI Calculation");
					}
				}
				cell.setCellStyle(ExcelStyleUtil.createNavyBlueStyle(workbook));
			}
			sheet.setColumnWidth(i, 256 * 5);
		}

		// Add merged regions for gray KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(2, 4, 1, 2));
		sheet.addMergedRegion(new CellRangeAddress(5, 37, 1, 1));
		sheet.addMergedRegion(new CellRangeAddress(5, 37, 2, 2));

		sheet.addMergedRegion(new CellRangeAddress(40, 76, 1, 1));
		sheet.addMergedRegion(new CellRangeAddress(40, 76, 2, 2));

		// Create text for gray KPI Calc box
		for (int i = 2; i < 76; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 1; j < 3; j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					cell = row.createCell(j);
				}
				if (i == 2 && j == 1) {
					cell.setCellValue("Fill rate - % of jobs filled vs job received\n"
							+ "(exclusive) (non-exclusive)\n"
							+ "**excluding cancelled reqs"); 
				}
				else if (i == 5 && j == 1) { 
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE # EXCLUSIVE JOBS FILLED\nFilter from the \"Closed\" tab:\n" +
									"1) Column N (Vendor) ->\n filter by vendor name \n(ex: only select JTS)\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"3) Column AM (Filled by) ->\nleave only your vendor selected\n" +
									"4) Enter 3 of reqs filled in the Quarter to KPI calculation\n" +
									"**(ex: per filtered data, if Entech\nfilled 15 exclusive positions in Q3, #\n jobs filled = 15)\n" 
							);
					richString.applyFont(0, 37, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(38, 67, ExcelStyleUtil.createBoldStyle(workbook));
					cell.setCellValue(richString);
				}
				else if (i == 5 && j == 2) { 
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE # EXCLUSIVE JOBS REC'D\nFilter from the \"Closed\" tab:\n" + 
									"1) Column J (Opened) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"2) Column N (Vendor) ->\n filter by vendor name \n(ex: only select JTS)\n" +
									"3) Column AI (Close Date) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"4) Make note of how many exclusive jobs received in quarter were filled. You'll add this number determined in the 'Open' tab calculation below.\n" + 
									"\nFilter from the \"Open\" tab:\n" + 
									"1) Column J (Opened) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"2) Column N (Vendor) ->\n filter by vendor name \n(ex: only select JTS)\n" +
									"3) If not exlusive jobs remain open, add\nthe filtered data from the closed tab to\nthe KPI calculation. If there are jobs still\nopen, add those to your 'Closed' tab\ndata and add to KPI calculation." 
							);
					richString.applyFont(0, 37, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(38, 67, ExcelStyleUtil.createBoldStyle(workbook));
					richString.applyFont(504, 532, ExcelStyleUtil.createBoldStyle(workbook));
					cell.setCellValue(richString);
				}
				else if (i == 40 && j == 1) { 
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE # NON-EXCL JOBS FILLED\nFilter from the \"Closed\" tab:\n" +
									"1) Column N (Vendor) ->\n filter by reqs open to ALL and remove\nreqs assigned to individual vendors\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"3) Column AM (Filled by) ->\nleave only your vendor selected\n" +
									"4) Enter 3 of reqs filled in the Quarter to KPI calculation\n" +
									"**(ex: per filtered data, if Entech\nfilled 40 non-exclusive positions in\n Q3, # jobs filled = 40)\n" 
							);
					richString.applyFont(0, 37, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(38, 67, ExcelStyleUtil.createBoldStyle(workbook));
					cell.setCellValue(richString);
				}
				else if (i == 40 && j == 2) { 
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE # NON-EXCL JOBS FILLED\nFilter from the \"Closed\" tab:\n" +
									"1) Column J (Opened) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"2) Column N (Vendor) ->\n filter by reqs open to ALL and remove\n reqs assigned to individual vendors\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"3) Column AM (Filled by) ->\nleave only your vendor selected\n" +
									"4) Make note of how many non-exclusive\n jobs receive in quarter were filled.\nYou'll add this number determined in\n the 'Open' tab calculation below.\n" +
									"\nFilter from the \"Open\" tab:\n" + 
									"1) Column J (Opened) ->\n filter by SLA quarter (ex: unselect all\n months except July, Aug, Sept for Q3\n SLA\n" +
									"2) Column N (Vendor) ->\n filter by reqs open to ALL and remove\nreqs assigned to individual vendors\n" +
									"3) If there are OTA jobs still open, add\nthose to your 'Closed' tab and add\nto KPI calculation. (ex: vendor partners\nclosed 150 jobs in Q3, but you\ndetermined 23 remain open. Total\nnon-exclusive jobs received = 173)" 
							);
					richString.applyFont(0, 37, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(38, 67, ExcelStyleUtil.createBoldStyle(workbook));
					richString.applyFont(597, 626, ExcelStyleUtil.createBoldStyle(workbook));
					cell.setCellValue(richString);
				}
				sheet.setColumnWidth(j, 256 * 38);
			}
		}

		// set style for gray KPI Calc box
		CellRangeAddress mergedRegion = new CellRangeAddress(2, 76, 1, 2);
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
		sheet.addMergedRegion(new CellRangeAddress(0, 2, 6, 7));
		sheet.addMergedRegion(new CellRangeAddress(3, 3, 6, 7));
		sheet.addMergedRegion(new CellRangeAddress(7, 8, 6, 7));
		sheet.addMergedRegion(new CellRangeAddress(9, 9, 6, 7));

		// set style for ratio of resumes to interviews table
		for(int r = 0; r < 13; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				row = sheet.createRow(r);
			}
			Cell cell = row.createCell(6);
			if(r == 0) {
				cell.setCellValue("Fill Rate\nKPI Calculation");
				cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
			}
			else if(r == 3) {
				cell.setCellValue("EXCLUSIVE JOBS");
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
			}
			else if(r == 4 || r == 10) {
				cell.setCellValue("# of jobs filled");
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(r == 4 ? jobsFilledExc : jobsFilledNon);
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
			else if(r == 5 || r == 11) {
				cell.setCellValue("# of jobs received");
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(r == 5 ? jobsReceivedExc : jobsReceivedNon);
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));			}
			else if(r == 6 || r == 12) {
				cell.setCellValue("Ratio:");
				cell.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
				Cell cell2 = row.createCell(7);
				String calculateExcKPI = calculateExcKPI(quarter, year, filePath);
				String calculateNonKPI = calculateNonKPI(quarter, year, filePath);
				cell2.setCellValue(r == 6 ? calculateExcKPI : calculateNonKPI);
				cell2.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
				Cell cell3 = row.createCell(8);
				if(r == 6) {
					cell3.setCellValue(" = " + jobsFilledExc + " / " + jobsReceivedExc + " * 100");
				} else {
					cell3.setCellValue(" = " + jobsFilledNon + " / " + jobsReceivedNon + " * 100");
				}
				cell3.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
			else if(r == 7) {
				cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
			}
			else if(r == 9) {
				cell.setCellValue("NON-EXCLUSIVE JOBS");
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
			}
		}
		sheet.setColumnWidth(6, 256 * 20);
		sheet.setColumnWidth(7, 256 * 20);
		sheet.setColumnWidth(8, 256 * 15);
	}
}
