package main;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

public class QuarterCheck {
	// get date from cell
	public static Date getDateFromCell(Cell cell) {
		if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
			return cell.getDateCellValue();
		} else if (cell.getCellType() == CellType.STRING) {
			String dateStr = cell.getStringCellValue();
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy", Locale.ENGLISH);
			try {
				return sdf.parse(dateStr);
			} catch (ParseException e) {
				e.printStackTrace();
			}
		}
		return null;
	}

	// check if date is in correct quarter
	public static boolean isInCorrectQuarter(Date date, int quarter, int year) {
		List<String> monthsInQuarter = getMonthsInQuarter(quarter);
		boolean result = isDateInQuarter(date, monthsInQuarter, year);
		return result;
	}

	// get months in quarter as an Array
	private static List<String> getMonthsInQuarter(int quarter) {
		switch (quarter) {
		case 1:
			return Arrays.asList("January", "February", "March");
		case 2:
			return Arrays.asList("April", "May", "June");
		case 3:
			return Arrays.asList("July", "August", "September");
		case 4:
			return Arrays.asList("October", "November", "December");
		default:
			throw new IllegalArgumentException("Invalid quarter: " + quarter);
		}
	}

	// check if date is in quarter
	private static boolean isDateInQuarter(Date date, List<String> monthsInQuarter, int year) {
		SimpleDateFormat monthFormat = new SimpleDateFormat("MMMM", Locale.ENGLISH);
		SimpleDateFormat yearFormat = new SimpleDateFormat("yyyy", Locale.ENGLISH);

		String month = monthFormat.format(date);
		int dateYear = Integer.parseInt(yearFormat.format(date));

		return monthsInQuarter.contains(month) && dateYear == year;
	}

	// check if date is in future
	public static boolean isDateInFuture(Date date, int quarter, int year) {
		SimpleDateFormat monthFormat = new SimpleDateFormat("MMMM", Locale.ENGLISH);
		SimpleDateFormat yearFormat = new SimpleDateFormat("yyyy", Locale.ENGLISH);

		String month = monthFormat.format(date);
		int dateYear = Integer.parseInt(yearFormat.format(date));
		int dateQuarter = getQuarter(date);

		if (dateYear > year) {
			return true;
		} else if (dateYear == year && dateQuarter > quarter) {
			return true;
		}

		return false;
	}
	
	// check if date is in future
	public static boolean isDateInPast(Date date, int quarter, int year) {
		SimpleDateFormat monthFormat = new SimpleDateFormat("MMMM", Locale.ENGLISH);
		SimpleDateFormat yearFormat = new SimpleDateFormat("yyyy", Locale.ENGLISH);

		String month = monthFormat.format(date);
		int dateYear = Integer.parseInt(yearFormat.format(date));
		int dateQuarter = getQuarter(date);

		if (dateYear < year) {
			return true;
		} else if (dateYear == year && dateQuarter < quarter) {
			return true;
		}

		return false;
	}
	
	private static int getQuarter(Date date) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		int month = cal.get(Calendar.MONTH) + 1;
		return (month + 2) / 3;
	}


}

