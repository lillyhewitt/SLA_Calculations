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

public class Failed_Hires_Report {
	// create variables to store KPI, headcount, and failed hires
	public static String KPI = "";
	public static String KPIaccepted = "";
	public static int headCount = 0;
	public static int failedHires = 0;
	public static int successfulHires = 0;

	// Read from an Excel file
	public static void read_From_Excel(String fileName, int quarter, int year) {
		try {
			// Creating a new file instance
			File file = new File(fileName);
			// Obtaining bytes from the file
			FileInputStream fis = new FileInputStream(file);
			// Creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);

			// loop through cells in sheet 1 (Roster)
			XSSFSheet sheet = wb.getSheetAt(0);

			// keep track of which table is data is being read from
			int VGIcount = 0;

			// index through rows in the sheet
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					break; // Skip if the row is null
				}

				Cell rcCell = row.getCell(3);  // Column D
				Cell titleCell = row.getCell(0);  // Column A

				// make sure data is only read from first 2 tables
				if (rcCell != null && rcCell.getCellType() == CellType.STRING) {
					if (titleCell != null && titleCell.getCellType() == CellType.STRING) {
						String cellValue = titleCell.getStringCellValue().replace("\n", " ");
						if (cellValue.equals("VGI Crew ID")) {
							VGIcount++;
							if(VGIcount == 2) {
								rowIndex = sheet.getLastRowNum();
							}
						}
					}

					// process row if it is not a header row 
					else if (!rcCell.getStringCellValue().equals("Name")) {
						processRow(row, quarter, year);
					}
				}
			}

			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Find head count and num of failed hires this quarter by checking dates of each entry in the roster and look for "fail" in notes
	private static void processRow(Row row, int quarter, int year) {
		Cell endCell = row.getCell(17);  // Column R
		Cell startCell = row.getCell(16);  // Column Q
		Cell notesCell = row.getCell(19);  // Column T

		if(startCell != null && startCell.getCellType() == CellType.NUMERIC && endCell != null && endCell.getCellType() == CellType.NUMERIC) {
			Date startDate = QuarterCheck.getDateFromCell(startCell);
			Date endDate = QuarterCheck.getDateFromCell(endCell);
			// check if start date is in the future (after this quarter) or end date has past
			if (!QuarterCheck.isDateInFuture(startDate, quarter, year) && !QuarterCheck.isDateInPast(endDate, quarter, year)) {
				headCount++;
				// check if "fail" is included in notes section, increase fail hired num if so
				if(notesCell != null && notesCell.getCellType() == CellType.STRING && containsFail(notesCell)) {
					failedHires++;
				}
			}
		}
	}

	// Check if notes section contains "fail"
	private static boolean containsFail(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String cellValue = cell.getStringCellValue().toLowerCase();
			return cellValue.contains("fail");
		}
		return false;
	}

	public static String calculateKPI(int quarter, int year, String filePath) {
		// reset variable 
		KPI = "";

		// Calculate KPI
		if(headCount > 0) {
			double ratio = (double) failedHires / headCount * 100;
			String ratioFormatted = String.format("%.2f", ratio) + "%";
			KPI = ratioFormatted;
		}
		return KPI;
	}

	public static String calculateAcceptableKPI(int quarter, int year, String filePath) {
		// reset variables
		KPIaccepted = "";
		headCount = 0;
		successfulHires = 0;
		failedHires = 0;

		// Get # of failed hires and headcount
		read_From_Excel(filePath, quarter, year);

		// calculate KPI of successful hires
		successfulHires = headCount - failedHires;
		double ratio = (double) successfulHires / headCount * 100;
		String ratioFormatted = String.format("%.2f", ratio) + "%";
		KPIaccepted = ratioFormatted;

		return KPIaccepted;
	}

	// Write Ratio Resumes to Interviews Exc to Excel file
	public void write_Failed_Hires(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year, String filePath) {
		// Add merged regions for blue KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 3));
		sheet.addMergedRegion(new CellRangeAddress(23, 24, 0, 3));
		sheet.addMergedRegion(new CellRangeAddress(2, 22, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 22, 3, 3));

		// set title for KPI box and style
		for (int i = 0; i < 24; i++) {
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
		sheet.addMergedRegion(new CellRangeAddress(5, 22, 1, 1));
		sheet.addMergedRegion(new CellRangeAddress(5, 22, 2, 2));

		// Create text for gray KPI Calc box
		for (int i = 2; i < 23; i++) {
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
					cell.setCellValue("Number of failed hires"); 
				}
				else if (i == 5 && j == 1) { 
					XSSFRichTextString richString = new XSSFRichTextString(
							"TO DETERMINE # OF FAILED HIRES IN QUARTER\n" +
									"\nVendor company pulling SLA data should\n" + 
									"use their own attrition records to\n" +
									"determine the number of resources that did not\n" +
									"complete their assignments in the\n" +
									"quarter. This should include resources\n" +
									"who quit and/or were terminated for\n" +
									"negative reasons prior to anticipated\n" +
									"assignment completion date (anyone who\n" +
									"did not complete the anticipated duration\n" +
									"of the assignment. Do not include\n" +
									"tenured resources or those who\n" +
									"complete the anticipated duration). Once\n" +
									"determined, add that number to the KPI\n Calculation."
							);
					richString.applyFont(0, 43, ExcelStyleUtil.createUnderlineStyle(workbook));
					cell.setCellValue(richString);
				}
				else if (i == 5 && j == 2) { 
					XSSFRichTextString richString = new XSSFRichTextString(
							"TO DETERMINE TOTAL HEADCOUNT IN QUARTER\n" + 
									"\nVendor company pulling SLA data should\n" +
									"use their own records to determine their\n" +
									"headcount for the quarter. Once\n" +
									"determined, add that number to the KPI\n Calculation."
							);
					richString.applyFont(0, 41, ExcelStyleUtil.createUnderlineStyle(workbook));
					cell.setCellValue(richString);
				}
				sheet.setColumnWidth(j, 256 * 38);
			}
		}

		// set style for gray KPI Calc box
		CellRangeAddress mergedRegion = new CellRangeAddress(2, 22, 1, 2);
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

		// add text and style for ratio of resumes to interviews table
		for(int r = 0; r < 6; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				row = sheet.createRow(r);
			}
			Cell cell = row.createCell(6);
			if(r == 0) {
				cell.setCellValue("Number of Never Starts\nKPI Calculation");
				cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
			}
			else if(r == 3) {
				cell.setCellValue("# Failed hires");
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(failedHires);
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
			else if(r == 4) {
				cell.setCellValue("# quarter headcount");
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(headCount);
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
			else if(r == 5) {
				cell.setCellValue("Ratio: ");
				cell.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
				Cell cell2 = row.createCell(7);
				cell2.setCellValue(calculateKPI(quarter, year, filePath));
				cell2.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
			}
		}
		sheet.setColumnWidth(6, 256 * 20);
		sheet.setColumnWidth(7, 256 * 20);
	}
}
