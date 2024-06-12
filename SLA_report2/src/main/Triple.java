package main;

import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.text.SimpleDateFormat;

public class Triple {
	private Date date1;
	private Date date2;
	private double value;

	// constructor
	public Triple(Date date1, Date date2, double weeksOnHold) {
		this.date1 = date1;
		this.date2 = date2;
		this.value = weeksOnHold;
	}

	// format date to MM/dd/yyyy
	public String formatter(Date date) {
		SimpleDateFormat formatter = new SimpleDateFormat("MM/dd/yyyy");
		return formatter.format(date);
	}

	// getters and setters
	public Date getDate1() {
		return date1;
	}

	public void setDate1(Date date1) {
		this.date1 = date1;
	}

	public Date getDate2() {
		return date2;
	}

	public void setDate2(Date date2) {
		this.date2 = date2;
	}

	public double getValue() {
		return value;
	}

	public void setValue(int value) {
		this.value = value;
	}

	// calculate business days between two dates

	public int calculateBusinessDays() {
		if (date1 == null || date2 == null) {
			throw new IllegalArgumentException("Dates must not be null");
		}
		Calendar startCal = new GregorianCalendar();
		startCal.setTime(date1);
		Calendar endCal = new GregorianCalendar();
		endCal.setTime(date2);

		int workDays = 0;

		// Check if start and end are the same
		if (startCal.getTimeInMillis() == endCal.getTimeInMillis() && 
				startCal.get(Calendar.DAY_OF_WEEK) != Calendar.SATURDAY && 
				startCal.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY) {
			return 1;
		}

		if (startCal.getTimeInMillis() > endCal.getTimeInMillis()) {
			startCal.setTime(date2);
			endCal.setTime(date1);
		}

		do {
			if (startCal.get(Calendar.DAY_OF_WEEK) != Calendar.SATURDAY && startCal.get(Calendar.DAY_OF_WEEK) != Calendar.SUNDAY) {
				++workDays;
			}
			startCal.add(Calendar.DAY_OF_MONTH, 1);
		} while (startCal.getTimeInMillis() <= endCal.getTimeInMillis()); // including end date

		return workDays;
	}

}