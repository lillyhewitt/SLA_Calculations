package main;

import java.io.File;
import java.io.FileInputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

public class Resumes_Interview_Exc_Report {
	// Create variables to store KPI and resume/interview pairs
	public static String KPI = "";
	public static List<Pair<Integer, Integer>> list = new ArrayList<>();
	public static int countExc = 0;

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

					// Process the row based on the sheet type
					Cell dateCell = row.getCell(9);
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

	// Find number of submitted resumes and interviews conducted based on date, vendor, fill, status, and submitted resumes cell values
	private static void processRow(Row row, int quarter, int year, boolean isClosedSheet) {
		Cell dateCell = row.getCell(9);  // Column Column J 
		Cell vendorCell = row.getCell(13);  // Column N
		Cell fillCell = isClosedSheet ? row.getCell(37) : null;  // Column AL
		Cell statusCell = row.getCell(24);  // Column Y
		Cell submittedResCell = row.getCell(22);  // Column W
		int submitted = 0;
		int interviews = 0;

		// check if open date is in this quarter
		Date date = QuarterCheck.getDateFromCell(dateCell);
		if (date != null && QuarterCheck.isInCorrectQuarter(date, quarter, year)) {
			// check for just Entech as vendor
			if (VendorCheck.isJustVendor(vendorCell)) { 
				// If status is not cancelled and submittedResCell has a numeric value, increase the submitted count on Closed tab
				if (isClosedSheet && fillCell != null && !"C".equals(fillCell.getStringCellValue())) {
					if(submittedResCell != null && submittedResCell.getCellType() != CellType.STRING) {
						submitted = (int) submittedResCell.getNumericCellValue();
					}
					countExc++;
				}
				// If submittedResCell has numeric value, increase the submitted count on Open tab
				if(!isClosedSheet && submittedResCell != null) {
					submitted = (int) submittedResCell.getNumericCellValue();
					countExc++;
				}
				// Find number of "int" found in notes to get the number of interviews
				if (countInts(statusCell) > 0) {
					interviews = countInts(statusCell);
				}
			}
		}

		// Add the submitted and interviews count to the list
		if(submitted > 0) {
			list.add(Pair.create(submitted, interviews));
		}
	}

	// count amount of interviews conducted, only on lines starting with "e" and containing "int"
	private static int countInts(Cell cell) {
		int countInts = 0;
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String cellValue = cell.getStringCellValue().toLowerCase();

			// Define the pattern to match "e - characters int M/D"
	        Pattern pattern = Pattern.compile("e\\s*- .* int(v)?\\s*(\\d{1,2}/\\d{1,2})");
			Matcher matcher = pattern.matcher(cellValue);

			// If the pattern is found, add to count
			while (matcher.find()) {
				// Get the matched string
				String match = matcher.group();

				// Count each "int M/D" in the matched string
	            Matcher intMatcher = Pattern.compile("int(v)?\\s*\\d{1,2}/\\d{1,2}").matcher(match);
				while (intMatcher.find()) {
					countInts++;
				}
			}
		}
		return countInts;
	}

	// calculate KPI
	public static String calculateKPI(int quarter, int year, String filePath) {
		// reset variables
		KPI = "";
		list = new ArrayList<>();

		// Get active and submitted resumes
		read_From_Excel(filePath, quarter, year);

		int sumResumes = 0;
		int sumInterviews = 0;

		// Loop through the list to show the number of resumes and interviews per entry
		for (Pair<Integer, Integer> pair : list) {
			sumResumes += pair.getFirst();
			sumInterviews += pair.getSecond();
		}

		// Calculate KPI
		double ratio = ((double) sumInterviews/sumResumes) * 100;
		String ratioFormatted = String.format("%.2f", ratio) + "%";
		KPI = ratioFormatted;
		return KPI;
	}

	// Write Ratio Resumes to Interviews NonExc to Excel file
	public void write_Resumes_Interviews_Exc(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year, String filePath) {
		// Add merged regions for blue KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 4));
		sheet.addMergedRegion(new CellRangeAddress(45, 46, 0, 4));
		sheet.addMergedRegion(new CellRangeAddress(2, 44, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 44, 4, 4));

		// Set title for KPI box and style
		for (int i = 0; i < 47; i++) {
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
		sheet.addMergedRegion(new CellRangeAddress(5, 44, 1, 3));

		// Create text for gray KPI Calc box
		for (int i = 2; i < 45; i++) { 
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
					// Create a rich text string for the second row
					XSSFRichTextString richString = new XSSFRichTextString("Ratio of resumes to interviews \n**excluding cancelled reqs [Break \nthis separately for exclusive and non-exclusive]");
					// Apply the bold style to the specified characters
					richString.applyFont(59, 115, ExcelStyleUtil.createBoldStyle(workbook));
					// Set the cell value to the rich text string
					cell.setCellValue(richString);
					cell.setCellStyle(ExcelStyleUtil.createResumeHeader(workbook));
				} else if (i == 4){
					cell.setCellStyle(ExcelStyleUtil.createResumeHeader(workbook));
				}
				else if (i == 5) { // Correct the condition to apply to the correct row
					// Create a rich text string for the fifth row
					XSSFRichTextString richString = new XSSFRichTextString(
							"\nFilter from the \"Closed\" tab:\n" +
									"1) Column J (Opened) ->\n filter by SLA quarter (ex. unselect all \nmonths except July, Aug, Sept for Q3\n SLA)\n" +
									"2) Column N (Vendor) ->\n filter by \"All\" and vendor name \n(ex: All + JTS)\n" +
									"3) Column AL (Filled/Cancelled) ->\n unselect \'C\' (only filled reqs remain)\n" +
									"4) Copy vendor data in \"Resumes Active\" \nand \"Resumes Submitted\"\n columns and paste into formula on\n this worksheet **(ex: RT would select data starting below Row 2 in Columns V and W through the end of the data set)\n" +
									"\nFilter from the \"Open\" Tab:\n" +
									"1) Column J (Opened) ->\n filter by SLA quarter (ex. unselect all months except July, Aug, Sept for Q3 SLA\n" +
									"2) Column N (Vendor) -> filter by \"All\" and vendor name \n(ex: \"All\" + JTS)\n" +
									"3) Copy Vendor data in \"Resumes\n Active\" and \"Resumes Submitted\"\n columns and paste into formula on\n this worksheet under data from\n opened tab **(ex: RT would select data starting\n below Row 2 in Columns V and W\n through the end of the data set\n" +
									"\n**It is good practice to manually verify\n \"Resumes Active\" and \"Resumes Submitted\"\n are accurate against the\n Status column"
							);
					// Apply the bold style to the specified characters
					richString.applyFont(0, 31, ExcelStyleUtil.createBoldStyle(workbook));
					richString.applyFont(504, 533, ExcelStyleUtil.createBoldStyle(workbook));
					// Set the cell value to the rich text string
					cell.setCellValue(richString);
					cell.setCellStyle(ExcelStyleUtil.createResumeGray(workbook));
				}
				sheet.setColumnWidth(j, 256 * 15);
			}
		}

		// Add merged regions ratio of resumes to interviews table
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 8, 9));

		// Loop only for rows 0 and 2 to set the title for the ratio of resumes to interviews table
		for(int r = 0; r < 3; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				row = sheet.createRow(r);
			}
			if(r == 0) {
				Cell cell = row.createCell(8);
				cell.setCellValue("Ratio of resumes to\n interviews");
				cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
			}
			else if(r == 2) {
				Cell cell = row.createCell(8);
				cell.setCellValue("# of Resumes");
				cell.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
				Cell cell2 = row.createCell(9);
				cell2.setCellValue("# of Interviews");
				cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
			}
		}

		int sumResumes = 0;
		int sumInterviews = 0;

		int rNew = 3; // Start from the third row
		// Loop through the list to show the number of resumes and interviews per entry
		for (Pair<Integer, Integer> pair : list) {
			Row row = sheet.getRow(rNew);
			if (row == null) {
				row = sheet.createRow(rNew);
			}

			Cell cell = row.createCell(8);
			cell.setCellValue(pair.getFirst()); // Set the key in column 11
			sumResumes += pair.getFirst();
			cell.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

			Cell cell2 = row.createCell(9);
			cell2.setCellValue(pair.getSecond()); // Set the value in column 12
			sumInterviews += pair.getSecond();
			cell2.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

			rNew++; // Move to the next row
		}

		// Set column widths outside the loop
		sheet.setColumnWidth(8, 256 * 15);
		sheet.setColumnWidth(9, 256 * 15);

		// Add merged regions ratio of resumes to interviews table
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 11, 12));
		sheet.addMergedRegion(new CellRangeAddress(2, 3, 11, 12));

		// Create text and style for ratio of resumes to interviews table
		CellRangeAddress mergedRegion = new CellRangeAddress(2, 3, 11, 12);
		for (int r = 0; r < 4; r++) {
			Row row = sheet.getRow(r);
			if (row == null) {
				row = sheet.createRow(r);
			}
			if (r == 0) {
				Cell cell = row.createCell(11);
				cell.setCellValue("KPI Calculation");
				cell.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));
			} else if (r == 2) {
				Cell cell = row.createCell(11);
				String KPICalc = calculateKPI(quarter, year, filePath);
				cell.setCellValue("Ratio:   " + KPICalc + "\n" + sumResumes + " resumes / " + sumInterviews + " interviews");
				cell.setCellStyle(ExcelStyleUtil.createYellowStyle(workbook));
			}
		}
		sheet.setColumnWidth(11, 256 * 20);


		// Set style for ratio of resumes to interviews table
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
