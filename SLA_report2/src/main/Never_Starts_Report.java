package main;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Never_Starts_Report {
	// create variables to store kpi, acceptances and rescinded acceptances
	public static String KPI = "";
	public static List<Pair<Integer, Integer>> list = new ArrayList<>();
	public static int acceptances = 0;
	public static int rescinded = 0;

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

			// check if terminations entry is found
			boolean check = false;

			// find acceptances this quarter
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					System.out.println("Skipping row " + rowIndex);
					break; // Skip if the row is null
				}

				Cell rcCell = row.getCell(3);  // Column D
				Cell titleCell = row.getCell(0);  // Column A

				if (rcCell != null && rcCell.getCellType() == CellType.STRING) {
					// make sure data is only read from first table
					if (titleCell != null && titleCell.getCellType() == CellType.STRING) {
						String cellValue = titleCell.getStringCellValue().replace("\n", " ");
						if (cellValue.equals("VGI Crew ID")) {
							break;
						}
					}
					else {
						// check if start date is set in future
						Cell startCell = row.getCell(16);  // Column Q
						if(startCell != null && startCell.getCellType() == CellType.NUMERIC) {
							Date startDate = QuarterCheck.getDateFromCell(startCell);
							Cell vgiCell = row.getCell(0);
							// check if start date is in the future or starting in this quarter 
							if (QuarterCheck.isDateInFuture(startDate, quarter, year) || QuarterCheck.isInCorrectQuarter(startDate, quarter, year)) {
								acceptances++;
							}
						}
						// check if start date is blank
						else if (startCell == null) {
							acceptances++;
						}
					}
				}
			}

			// find rescinded acceptances this quarter
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row == null) {
					break; // Skip if the row is null
				}
				
				// only read from Rescinded Acceptance table
				if(check) {
					Cell startCell = row.getCell(16);  // Column Q

					if(startCell != null && startCell.getCellType() == CellType.NUMERIC) {
						Date startDate = QuarterCheck.getDateFromCell(startCell);
						// check if start date is in current quarter
						if (QuarterCheck.isInCorrectQuarter(startDate, quarter, year)) {
							rescinded++;
						}
					}
				}

				// find beginning of Rescinded Acceptance table
				Cell rescindedTitleCell = row.getCell(0);  // Column A
				if (rescindedTitleCell != null && rescindedTitleCell.getCellType() == CellType.STRING && rescindedTitleCell.getStringCellValue().equals("Rescinded Acceptance/Withdrawn Candidates")) {
					check = true;
					rowIndex++;
				}
				// find end of Rescinded Acceptance table
				else if (rescindedTitleCell != null && rescindedTitleCell.getCellType() == CellType.STRING
						&& rescindedTitleCell.getStringCellValue().equals("Resume Fraud")) {
					check = false;
				}
			}
			// add data to list 
			list.add(Pair.create(acceptances, rescinded));
			
			// calculate KPI
			double ratio = (double) rescinded / acceptances * 100;
			String ratioFormatted = String.format("%.2f", ratio) + "%";
			KPI = ratioFormatted;

			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static String calculateKPI(int quarter, int year, String filePath) {
		// reset variables
		KPI = "";
		list = new ArrayList<>();
		acceptances = 0;
		rescinded = 0;

		// Get never starts
		read_From_Excel(filePath, quarter, year);
		
		// Calculate KPI
		double ratio = (double) rescinded / acceptances * 100;
		String ratioFormatted = String.format("%.2f", ratio) + "%";
		KPI = ratioFormatted;
		return KPI;
	}

}
