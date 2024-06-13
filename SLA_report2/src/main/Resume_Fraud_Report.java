package main;

import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Resume_Fraud_Report {
	// create variables to store name and fraud report
	public static int KPI = 0;
	public static List<Pair<String, String>> list = new ArrayList<>();
	public static int fraudCount = 0;

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
			boolean check = false;

			// index through rows in the sheet
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					break; // Skip if the row is null
				}

				// read from resume fraud table
				if(check) {
					Cell lastNameCell = row.getCell(1);  // Column B
					Cell firstNameCell = row.getCell(2);  // Column C
					Cell reportFraudCell = row.getCell(16);  // Column D

					// parse data to find name and fraud report if it exists
					if (lastNameCell != null && firstNameCell != null && reportFraudCell != null && reportFraudCell.getCellType() == CellType.STRING && !containsOld(reportFraudCell)) {
						String fullName = firstNameCell.getStringCellValue() + " " + lastNameCell.getStringCellValue();
						String reportFraud = reportFraudCell.getStringCellValue();
						list.add(Pair.create(fullName, reportFraud));
						fraudCount++;
					}
				}
				Cell fraudCell = row.getCell(0);  // Column A
				// check if the table is the resume fraud table
				if (fraudCell != null && fraudCell.getCellType() == CellType.STRING && fraudCell.getStringCellValue().equals("Resume Fraud")) {
					check = true;
					rowIndex++; // skip table header
				}
			}
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Check if cell contains the word "old"
	private static boolean containsOld(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String cellValue = cell.getStringCellValue().toLowerCase();
			return cellValue.contains("old");
		}
		return false;
	}

	public static int calculateKPI(int quarter, int year, String filePath) {
		// reset variables
		KPI = 0;
		list = new ArrayList<>();
		fraudCount = 0;

		// collect fraud resumes
		read_From_Excel(filePath, quarter, year);
		
		// Calculate KPI
		int count = 0;
		for (Pair<String, String> pair : list) {
			count++;
		}
		return KPI = count;
	}

	// Write Ratio Resumes to Interviews Exc to Excel file
	public void write_Resume_Fraud(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year, String filePath) {
		calculateKPI(quarter, year, filePath);

		// Add merged regions for blue KPI Calc box
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 6));
		sheet.addMergedRegion(new CellRangeAddress(14, 15, 0, 6));
		sheet.addMergedRegion(new CellRangeAddress(2, 13, 0, 0));
		sheet.addMergedRegion(new CellRangeAddress(2, 13, 6, 6));

		// Set the style for the blue KPI Calc box
		for (int i = 0; i < 15; i++) {
			Row row = sheet.getRow(i);
			if (row == null) {
				row = sheet.createRow(i);
			}
			for (int j = 0; j < 7; j++) { 
				Cell cell = row.getCell(j);
				if (cell == null) {
					cell = row.createCell(j);
				}
				cell.setCellStyle(ExcelStyleUtil.createNavyBlueStyle(workbook));
			}
		}

		// Add merged regions for the gray Resume Fraud box
		sheet.addMergedRegion(new CellRangeAddress(2, 13, 1, 5));

		// Create the text for the gray Resume Fraud box
		for (int i = 2; i < 6; i++) {
			Row row = sheet.getRow(i);
			for (int j = 1; j < 4; j++) { 
				Cell cell = row.getCell(j);
				XSSFRichTextString richString = new XSSFRichTextString(
						"Resume Fraud\n" +
								"**This will need to be reported by each vendor\n" +
								"based on their own records and communication\n" +
								"with GRSC leadership**"
						);
				richString.applyFont(ExcelStyleUtil.createRedStyle(workbook));
				richString.applyFont(0, 13, ExcelStyleUtil.createUnderlineStyle(workbook));
				cell.setCellValue(richString);
			}
		}

		// Set style for the gray Resume Fraud box
		CellRangeAddress mergedRegion = new CellRangeAddress(2, 13, 1, 5);
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
				else {
					cell.setCellStyle(ExcelStyleUtil.createGrayResumeFraudStyle(workbook));
				}
			}
		}

		// Add merged regions for the Resume Fraud entries table
		sheet.addMergedRegion(new CellRangeAddress(0, 1, 8, 9));
		Row row = sheet.getRow(0);
		Cell cellTitle = row.createCell(8);
		cellTitle.setCellValue("Resume Fraud");
		cellTitle.setCellStyle(ExcelStyleUtil.createNavyTableStyle(workbook));

		// Add entries of resume fraud to the table
		int rNew = 2; 
		int count = 0;
		for (Pair<String, String> pair : list) {
			Row rowObj = sheet.getRow(rNew);
			if (rowObj == null) {
				rowObj = sheet.createRow(rNew);
			}

			Cell cellName = rowObj.createCell(8);
			cellName.setCellValue(pair.getFirst()); // Set the key in column 1
			cellName.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

			Cell cellFraud = rowObj.createCell(9);
			cellFraud.setCellValue(pair.getSecond()); // Set the value in column 2
			cellFraud.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));

			rNew++; // Move to the next row
			count++;
		}
		sheet.setColumnWidth(8, 256 * 30);
		sheet.setColumnWidth(9, 256 * 70);
	}
}
