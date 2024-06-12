package main;

import java.io.File;
import java.io.FileInputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Respond_With_Resume_Report {
	public static double KPIexc = 0;
	public static double KPInon = 0;
	public static int excCount = 0;
	public static int excDayCounter = 0;
	public static int nonCount = 0;
	public static int nonDayCounter = 0;

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
					Cell dateCell = row.getCell(9);  // Column J
					if(i == 1) {
						if (dateCell.getCellType() != CellType.NUMERIC || dateCell.getDateCellValue() == null) {
							rowIndex = sheet.getLastRowNum();
						} else {
							//System.out.println(dateCell.getDateCellValue());
							processRow(row, quarter, year, false);
						}
					}else if (dateCell != null && dateCell.getDateCellValue() != null) {
						processRow(row, quarter, year, true);
					}
				}
				KPIexc = (double) excDayCounter / excCount;
				KPInon = (double) nonDayCounter / nonCount;

				KPIexc = Double.parseDouble(String.format("%.1f", (double) excDayCounter / excCount));
				KPInon = Double.parseDouble(String.format("%.1f", (double) nonDayCounter / nonCount));
			}
		}catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void processRow(Row row, int quarter, int year, boolean isClosedSheet) {
		Cell dateCell = row.getCell(9);  // Column AI for closed sheet, Column J for open sheet
		Cell vendorCell = row.getCell(13);  // Column N
		Cell infoCell = row.getCell(24);  // Column Y

		Date date = QuarterCheck.getDateFromCell(dateCell);
		if (date != null && QuarterCheck.isInCorrectQuarter(date, quarter, year) && infoCell != null && infoCell.getCellType() == CellType.STRING) {
			String recDate = findRecDate(infoCell);
			String firstSub = findSubDate(infoCell);
			// exclusive 
			if (VendorCheck.isJustVendor(vendorCell)) {
				excCount++;
				if(recDate != null && firstSub != null) {
					excDayCounter += calculateBusinessDays(recDate, firstSub, year);
				} 
			}
			// non-exclusive
			else if (VendorCheck.isAllVendor(vendorCell)) {
				nonCount++;
				if(recDate != null && firstSub != null) {
					nonDayCounter += calculateBusinessDays(recDate, firstSub, year);
				}
			}
		}
	}

	private static String findRecDate(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String cellValue = cell.getStringCellValue().toLowerCase();

			// Define the pattern to match "Sent to All" or "Sent to Entech" followed by a date
			Pattern pattern = Pattern.compile("(sent to all|sent to entech)\\s*(\\d{1,2}/\\d{1,2})");
			Matcher matcher = pattern.matcher(cellValue);

			// If the pattern is found, return the date
			if (matcher.find()) {
				return matcher.group(2);
			}
		}
		return null;
	}


	private static String findSubDate(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String cellValue = cell.getStringCellValue().toLowerCase();

			// Define the pattern to match "E - " or "E-" followed by any name, then "sub" followed by a date on same line
			Pattern pattern = Pattern.compile("e\\s*- .* sub\\s*(\\d{1,2}/\\d{1,2})");
			Matcher matcher = pattern.matcher(cellValue);

			Pattern patternInt = Pattern.compile("e\\s*- .* sub\\s*(\\d{1,2}/\\d{1,2})(,\\s*int\\s*\\d{1,2}/\\d{1,2}|-\\s*int\\s*\\d{1,2}/\\d{1,2})");
			Matcher matcherInt = patternInt.matcher(cellValue);

			// If the pattern is found, return the date
			if (matcher.find()) {
				return matcher.group(1);
			} else if (matcherInt.find()) {
				return matcherInt.group(1);
			}
		}
		return null;
	}


	private static long calculateBusinessDays(String recDate, String firstSub, int year) {
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
		LocalDate startDate = LocalDate.parse(recDate + "/" + year, formatter);
		LocalDate endDate = LocalDate.parse(firstSub + "/"+ year, formatter);

		// If the month of firstSub is less than recDate, subtract one year from recDate
		if (endDate.getMonthValue() < startDate.getMonthValue()) {
			startDate = startDate.minusYears(1);
		}


		int businessDays = 0;

		for (LocalDate date = startDate; !date.isAfter(endDate); date = date.plusDays(1)) {
			DayOfWeek dayOfWeek = date.getDayOfWeek();
			if (dayOfWeek != DayOfWeek.SATURDAY && dayOfWeek != DayOfWeek.SUNDAY) {
				businessDays++;
			}
		}

		return businessDays;
	}

	public static double calculateExcKPI(int quarter, int year, String filePath) {
		KPIexc = 0;
		KPInon = 0;
		excCount = 0;
		excDayCounter = 0;
		nonCount = 0;
		nonDayCounter = 0;

		// Get num of days between received job opening and first submission
		read_From_Excel(filePath, quarter, year);

		// Calculate KPI
		KPIexc = (double) excDayCounter / excCount;

		System.out.println("EXC KPI: " + KPIexc);
		return KPIexc;
	}

	public static double calculateNonKPI(int quarter, int year, String filePath) {
		// Calculate KPI
		KPInon = (double) nonDayCounter / nonCount;
		System.out.println("Non KPI: " + KPInon);
		return KPInon;
	}

}
