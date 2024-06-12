package main;

import java.io.File;
import java.io.FileInputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class Time_Accept_Report {
	// create variables to store KPI and data of important entries
	public static int KPIexc = 0;
	public static int KPInon = 0;
	public static List<Triple> listExc = new ArrayList<>();
	public static List<Triple> listNon = new ArrayList<>();
	public static double fillsExc = 0;
	public static double daysExc = 0;
	public static double fillsNon = 0;
	public static double daysNon = 0;

	// Read from an Excel file
	public static void read_From_Excel(String fileName, int quarter, int year) {
		try {
			// Creating a new file instance
			File file = new File(fileName);
			// Obtaining bytes from the file
			FileInputStream fis = new FileInputStream(file);
			// Creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);

			// loop through cells in sheet 2 (Closed)
			XSSFSheet sheet = wb.getSheetAt(2);

			// index through rows in the sheet
			for (int rowIndex = 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					break; // Skip if the row is null
				}
				Cell dateCell = row.getCell(9);  // Column AI
				if (dateCell != null) {
					processRow(row, quarter, year);
				}
			}
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Store data based on closed date, filled date, and vendor
	private static void processRow(Row row, int quarter, int year) {
		Cell vendorCell = row.getCell(13);  // Column N
		Cell closeDate = row.getCell(34);  // Column AI
		Cell fillDate = row.getCell(38);  // Column AM

		Date reqReceiveDate = null;
		Date reqCloseDate = null;
		double weeksOnHold = 0;

		Date date = QuarterCheck.getDateFromCell(closeDate);
		if (date != null && QuarterCheck.isInCorrectQuarter(date, quarter, year)) {
			// Check if vendor is only Entech and filled by Entech
			if (VendorCheck.isJustVendor(vendorCell) && "E".equals(fillDate.getStringCellValue())) {
				// store received and closed date of position
				reqReceiveDate = QuarterCheck.getDateFromCell(row.getCell(31));  // Column AF
				reqCloseDate = QuarterCheck.getDateFromCell(closeDate);
				// store number of weeks on hold
				if (row.getCell(28) != null) {
					weeksOnHold = row.getCell(28).getNumericCellValue(); // Column AC
				}
				else {
					weeksOnHold = 0;
				}
				// add stored data to exclusive list
				if(reqReceiveDate != null && reqCloseDate != null) {
					listExc.add(new Triple(reqReceiveDate, reqCloseDate, weeksOnHold));
				}
			}
			// check if vendor is All and filled by Entech
			else if (VendorCheck.isAllVendor(vendorCell) && "E".equals(fillDate.getStringCellValue())) {
				// store received and closed date of position
				reqReceiveDate = QuarterCheck.getDateFromCell(row.getCell(31));  // Column AF
				reqCloseDate = QuarterCheck.getDateFromCell(closeDate);
				// store number of weeks on hold
				if (row.getCell(28) != null) {
					weeksOnHold = (int) row.getCell(28).getNumericCellValue(); // Column AC
				}
				else {
					weeksOnHold = 0;
				}
				// add stored data to nonexclusive list
				if(reqReceiveDate != null && reqCloseDate != null) {
					listNon.add(new Triple(reqReceiveDate, reqCloseDate, weeksOnHold));
				}
			}
		}
	}

	public static int calculateExcKPI(int quarter, int year, String filePath) {
		KPIexc = 0;
		listExc = new ArrayList<>();
		KPInon = 0;
		listNon = new ArrayList<>();

		// Get time took to accept position
		read_From_Excel(filePath, quarter, year);

		// Calculate KPI
		fillsExc = 0;
		daysExc = 0;
		// Create 2 tables for exclusive and non-exclusive
		for(int r = 0; r < 309; r++) {
			if(r > 3) {
				int listIndex = r - 4; 
				List<Triple> currentList = listExc;
				if (listIndex < currentList.size()) { // Check if there is a Triple at this index
					Triple triple = currentList.get(listIndex); // Get the Triple object
					int businessDays = triple.calculateBusinessDays();
					double weeksOnHold = triple.getValue();
					double daysRounded = (double) Math.round(businessDays);
					double days  = daysRounded - (weeksOnHold*7);
					fillsExc++;
					daysExc += days;
				}
			}

		}
		double avg =  daysExc / fillsExc;
		KPIexc = (int) avg;

		return KPIexc;
	}

	public static int calculateNonKPI(int quarter, int year, String filePath) {
		// Calculate KPI
		fillsNon = 0;
		daysNon = 0;

		for(int r = 0; r < 309; r++) {
			if(r > 3) {
				int listIndex = r - 4; 
				List<Triple> currentList = listNon;
				if (listIndex < currentList.size()) { // Check if there is a Triple at this index
					Triple triple = currentList.get(listIndex); // Get the Triple object
					int businessDays = triple.calculateBusinessDays();
					double weeksOnHold = triple.getValue();
					double days  = businessDays - (weeksOnHold*7);
					fillsNon++;
					daysNon += days;
				}
			}
		}
		double avg = daysNon / fillsNon;
		KPInon = (int) avg;
		return KPInon;
	}

	// Write Ratio Resumes to Interviews NonExc to Excel file
	public void write_Time_to_Accept(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year, String filePath) {
		// Add merged regions for blue KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 4));
		sheet.addMergedRegion(new CellRangeAddress(33, 34, 1, 3));
		sheet.addMergedRegion(new CellRangeAddress(2, 63, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 63, 4, 4));
		sheet.addMergedRegion(new CellRangeAddress(64, 65, 0, 4));

		// set title for KPI box and style
		for (int i = 0; i < 66; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 0; j < 5; j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					cell = row.createCell(j);
					if (i == 0 && j == 0) {
						cell.setCellValue("KPI Calculation");
					}
				}
				cell.setCellStyle(ExcelStyleUtil.createNavyBlueStyle(workbook));
				sheet.setColumnWidth(j, 256 * 5);
			}
		}

		// Add merged regions for gray KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(2, 4, 1, 3));
		sheet.addMergedRegion(new CellRangeAddress(5, 32, 1, 3));
		sheet.addMergedRegion(new CellRangeAddress(35, 63, 1, 3));

		// Create text for gray KPI Calc box
		for (int i = 2; i < 64; i++) { 
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 1; j < 4; j++) {
				Cell cell = row.getCell(j);
				if (cell == null) {
					cell = row.createCell(j);
				}
				if (i == 2) {
					cell.setCellValue("Time to Accept/to Position Fulfillment\n"
							+ "(exclusive) (non-exclusive)");
					cell.setCellStyle(ExcelStyleUtil.createGrayHeaderStyle(workbook));
				} else if (i == 3 || i == 4){
					cell.setCellStyle(ExcelStyleUtil.createGrayHeaderStyle(workbook));
				}
				else if (i == 5) { // Correct the condition to apply to the correct row
					// Create a rich text string for the fifth row
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE TIME TO ACCEPT\nEXCLUSIVE JOBS IN QUARTER\nFilter from the \"Closed\" tab:\n" +
									"1) Column N (Vendor) ->\n filter by vendor name \n(ex: JTS only)\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex. unselect all\n months except July, Aug, Sept for Q3 SLA\n" +  
									"3) Column AM (Filled by) ->\nleave only your vendor selected" +
									"4) Copy the dates vendor received reqs\nfrom Req report (will be columns\nAE-AH, depending on vendor) into\nColumn I(Date rec'd) on this sheet.\n" +
									"**(Ex: RT would copy data in column\nAH starting from under row 2)" +
									"5) Copy the dates vendor closed the\n reqs from Req report (will be column\n AI, Closed Date) into Column J (Date\n closed) on this sheet." +
									"**(Ex: RT would copy data in\n column AI starting from under row 2)\n" +
									"6) Copy the weeks on hold (column\nAC) starting from under row 2 on the\n Req report to Column J (weeks on\n hold) on this sheet."                        
							);
					// Apply the bold style to the specified characters
					richString.applyFont(0, 54, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(55, 84, ExcelStyleUtil.createBoldStyle(workbook));
					// Set the cell value to the rich text string
					cell.setCellValue(richString);
					cell.setCellStyle(ExcelStyleUtil.createGrayStyle(workbook));
				}
				else if (i == 35) { // Correct the condition to apply to the correct row
					// Create a rich text string for the fifth row
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nTO DETERMINE TIME TO ACCEPT\nNON-EXCLUSIVE JOBS IN QUARTER\nFilter from the \"Closed\" tab:\n" +
									"1) Column N (Vendor) ->\n filter by vendor name \n(ex: JTS only)\n" +
									"2) Column AI (Close Date) ->\n filter by SLA quarter (ex. unselect all\n months except July, Aug, Sept for Q3 SLA\n" +  
									"3) Column AM (Filled by) ->\nleave only your vendor selected\n" +
									"4) Copy the dates vendor received reqs\nfrom Req report (will be columns\nAE-AH, depending on vendor) into\nColumn I(Date rec'd) on this sheet.\n" +
									"**(Ex: RT would copy data in column\nAH starting from under row 2)\n" +
									"5) Copy the dates vendor closed the\n reqs from Req report (will be column\n AI, Closed Date) into Column J (Date\n closed) on this sheet.\n" +
									"**(Ex: RT would copy data in\n column AI starting from under row 2)\n" +
									"6) Copy the weeks on hold (column\nAC) starting from under row 2 on the\n Req report to Column S (weeks on\n hold) on this sheet."                        
							);
					// Apply the bold style to the specified characters
					richString.applyFont(0, 59, ExcelStyleUtil.createUnderlineStyle(workbook));
					richString.applyFont(60, 89, ExcelStyleUtil.createBoldStyle(workbook));
					// Set the cell value to the rich text string
					cell.setCellValue(richString);
					cell.setCellStyle(ExcelStyleUtil.createGrayStyle(workbook));
				}
				else if(i != 33 && i != 34){
					cell.setCellStyle(ExcelStyleUtil.createGrayStyle(workbook));
				}
				sheet.setColumnWidth(j, 256 * 15);
			}
		}

		// Create 2 tables for exclusive and non-exclusive
		for(int interval = 0; interval < 2; interval++) {
			// Add merged regions ratio of resumes to interviews table
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 6+(interval*8), 9+(interval*8)));
			sheet.addMergedRegion(new CellRangeAddress(2, 2, 6+(interval*8), 9+(interval*8)));

			// set titles for position entry information
			for(int r = 0; r < 309; r++) {
				Row row = sheet.getRow(r);
				if (row == null) {
					row = sheet.createRow(r);
				}
				if(r == 0) {
					Cell cell = row.createCell(6+(interval*8));
					cell.setCellValue("Time to Accept/Position Fulfillment");
					cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
				}
				else if(r == 2) {
					Cell cell = row.createCell(6+(interval*8));
					if(interval == 0) {
						cell.setCellValue("EXCLUSIVE");
					}
					else {
						cell.setCellValue("NON-EXCLUSIVE");
					}
					cell.setCellStyle(ExcelStyleUtil.createPlainTableHeaderStyle(workbook));
				}
				else if(r == 3) {
					Cell cell = row.createCell(6+(interval*8));
					cell.setCellValue("Date rec'd");
					cell.setCellStyle(ExcelStyleUtil.createPlainTableHeaderStyle(workbook));
					Cell cell2 = row.createCell(7+(interval*8));
					cell2.setCellValue("Date closed");
					cell2.setCellStyle(ExcelStyleUtil.createPlainTableHeaderStyle(workbook));
					Cell cell3 = row.createCell(8+(interval*8));
					cell3.setCellValue("weeks on hold");
					cell3.setCellStyle(ExcelStyleUtil.createPlainTableHeaderStyle(workbook));
					Cell cell4 = row.createCell(9+(interval*8));
					cell4.setCellValue("Business Days");
					cell4.setCellStyle(ExcelStyleUtil.createPlainTableHeaderStyle(workbook));
				}
				else if(r > 3) {
					int listIndex = r - 4; // Adjust index to match list size starting from 0
					List<Triple> currentList = (interval == 0) ? listExc : listNon;
					if (listIndex < currentList.size()) { // Check if there is a Triple at this index
						Triple triple = currentList.get(listIndex); // Get the Triple object
						int businessDays = triple.calculateBusinessDays();
						if(interval == 0) {
							KPIexc = calculateExcKPI(quarter, year, filePath);
						} else {
							KPInon = calculateNonKPI(quarter, year, filePath);
						}

						Cell cell = row.createCell(6+(interval*8));
						cell.setCellValue(triple.formatter(triple.getDate1())); // Set the cell value to the first date
						cell.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

						Cell cell2 = row.createCell(7+(interval*8));
						cell2.setCellValue(triple.formatter(triple.getDate2())); // Set the cell value to the second date
						cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

						double weeksOnHold = triple.getValue();
						Cell cell3 = row.createCell(8+(interval*8));
						cell3.setCellValue(weeksOnHold); // Set the cell value to the integer value
						cell3.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

						Cell cell4 = row.createCell(9+(interval*8));
						double days  = businessDays - (weeksOnHold*7);
						double daysRounded = (double) Math.round(days);
						cell4.setCellValue(daysRounded); // Set the cell value to the calculated business days
						cell4.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
					}
				}
			}

			// set style for exlusive and nonexclusive tables
			CellRangeAddress mergedRegionTitle = new CellRangeAddress(2, 2, 6+(interval*8), 9+(interval*8));
			for (int rowNum = mergedRegionTitle.getFirstRow(); rowNum <= mergedRegionTitle.getLastRow(); rowNum++) {
				Row row = sheet.getRow(rowNum);
				if (row == null) {
					row = sheet.createRow(rowNum);
				}
				for (int colNum = mergedRegionTitle.getFirstColumn(); colNum <= mergedRegionTitle.getLastColumn(); colNum++) {
					Cell cell = row.getCell(colNum);
					if (cell == null) {
						cell = row.createCell(colNum);
					}
					cell.setCellStyle(ExcelStyleUtil.createPlainTableHeaderStyle(workbook));
				}
			}

			// Set column widths outside the loop
			sheet.setColumnWidth(6+(interval*8), 256 * 20);
			sheet.setColumnWidth(7+(interval*8), 256 * 20);
			sheet.setColumnWidth(8+(interval*8), 256 * 20);
			sheet.setColumnWidth(9+(interval*8), 256 * 20);

			// Add merged regions ratio of resumes to interviews table
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 11+(interval*8), 12+(interval*8)));
			sheet.addMergedRegion(new CellRangeAddress(2, 3, 11+(interval*8), 12+(interval*8)));

			// Create text and style for KPI Calculation box
			CellRangeAddress mergedRegion = new CellRangeAddress(2, 3, 11+(interval*8), 12+(interval*8));
			for (int r = 0; r < 4; r++) {
				Row row = sheet.getRow(r);
				if (row == null) {
					row = sheet.createRow(r);
				}
				if (r == 0) {
					Cell cell = row.createCell(11+(interval*8));
					cell.setCellValue("KPI Calculation");
					cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
				} else if (r == 2) {
					Cell cell = row.createCell(11+(interval*8));
					int avg = 0;
					if(interval == 0) {
						//	int excKPI = calculateExcKPI(quarter, year, filePath);
						cell.setCellValue("Days to " + KPIexc + "\n" + (int) fillsExc + " fills / " + (int) daysExc + " days");
					}
					else {
						//	int nonKPI = calculateNonKPI(quarter, year, filePath);
						cell.setCellValue("Days to " + KPInon + "\n" + (int) fillsNon + " fills / " + (int) daysNon + " days");
					}
					cell.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
				}
			}

			// Set style for KPI calculation box
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
					cell.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
				}
			}
		}
	}
}
