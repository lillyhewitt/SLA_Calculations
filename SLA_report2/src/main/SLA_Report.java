package main;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SLA_Report {
	// Write SLA Report to Excel file
	public void write_SLA_Report(XSSFWorkbook workbook, XSSFSheet sheet, HashMap<XSSFSheet, String> allSheets, int quarter, int year, String fileReqPath, String fileEntechPath) {
		try {
			// Add title to SLA report
			Row r0 = sheet.createRow(0);
			r0.setHeightInPoints(50); // Set row height to accommodate wrapped text
			Cell c0 = r0.createCell(0);
			c0.setCellValue("Entech - Q" + quarter + " SLAs \nProgram Level");
			c0.setCellStyle(ExcelStyleUtil.createNavyBlueStyle(workbook));

			// Create an array to store the quarters that need to be run
			int[] quarters = new int[3];
			int[] newYears = new int[3];
			int newYear = year - 1;
			if(quarter == 1) {	// quarter 1
				quarters = new int[] {4, 3, 2};
				newYears = new int[] {newYear, newYear, newYear};
			}
			else if(quarter == 2) { // quarter 2
				quarters = new int[] {1, 4, 3};
				newYears = new int[] {year, newYear, newYear};
			} else if (quarter == 3) { // quarter 3
				quarters = new int[] {2, 1, 4};
				newYears = new int[] {year, year, newYear};
			} else { // quarter 4
				quarters = new int[] {3, 2, 1};
				newYears = new int[] {year, year, year};
			}

			// Add Column Header row
			Row r1 = sheet.createRow(2);
			// Add Column headers to SLA report
			ArrayList<String> columnHeaders = new ArrayList<>(Arrays.asList(" ", "Performance Theme", "Theme Weight", "KPI", "Description", "KPI%", "Vendor Comments", " ", " ", "Last KPI %", "2nd last KPI %", "3rd last KPI %", " ", "Target", "Green", "Yellow", "Red", " "));
			for (int col = 0; col < columnHeaders.size(); col++) {
				Cell c = r1.createCell(col);
				String headerValue = columnHeaders.get(col);
				if (col == 0 || col == 7 || col == 12 || col == 17) {
					// Apply gray style to specific cells
					c.setCellStyle(ExcelStyleUtil.createBorderGrayStyle(workbook));
				} else {
					if (headerValue.equals("Last KPI %")) {
						c.setCellValue("Q" + quarters[0] + " " + newYears[0] + " KPI %");
					}
					else if (headerValue.equals("2nd last KPI %")) {
						c.setCellValue("Q" + quarters[1] + " " + newYears[1] + " KPI %");
					} else if (headerValue.equals("3rd last KPI %")) {
						c.setCellValue("Q" + quarters[2] + " " + newYears[2] + " KPI %");
					}
					else {
						c.setCellValue(headerValue);
					}
					// Apply border style to other cells
					c.setCellStyle(ExcelStyleUtil.createHeaderStyle(workbook));
				}
			}

			// Set column widths
			for (int col = 0; col < columnHeaders.size(); col++) {
				if (col == 0 || col == 7 ||col == 8 || col == 12 || col == 17) {
					// Set width of gray columns to 10 units
					sheet.setColumnWidth(col, 256 * 5);
				} else {
					// Set width of other columns to fit content (adjust as necessary)
					sheet.setColumnWidth(col, 4000);
				}
			}

			// Add merged regions to SLA report
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 17));
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 18, 21));
			sheet.addMergedRegion(new CellRangeAddress(2, 17, 21, 21));
			sheet.addMergedRegion(new CellRangeAddress(2, 17, 0, 0));
			sheet.addMergedRegion(new CellRangeAddress(2, 17, 7, 8));
			sheet.addMergedRegion(new CellRangeAddress(2, 17, 12, 12));
			sheet.addMergedRegion(new CellRangeAddress(2, 17, 17, 17));
			sheet.addMergedRegion(new CellRangeAddress(3, 17, 1, 1));
			sheet.addMergedRegion(new CellRangeAddress(3, 17, 2, 2));
			sheet.addMergedRegion(new CellRangeAddress(3, 8, 3, 3));
			sheet.addMergedRegion(new CellRangeAddress(9, 12, 3, 3));

			// Add column entires to SLA table
			ArrayList<String> kpiHeaders = new ArrayList<>(Arrays.asList("Interview to Hire Ratio (Timely Delivery of Resources", "Time to Fill Ratio", "Contractor Performance", "Failed Hires", "Assignment Completion Rate", "Resume Fraud/Proxy or Imposter Canidadate", "Provider Workforce Turbulence (Attrition)"));
			ArrayList<String> kpiDescriptions = new ArrayList<>(Arrays.asList("Respond with Resume (exclusive) in 2 weeks(open to all after 2 weeks)", "Respond with Resume(non-exclusive)", "Ratio of resumes to interviews(exclusive)", "Ratio of resumes to interviews(non-exclusive)", "Number of Offers declined (exclusive and non-exclusive) Track quarterly - on a trial basis ", "Number of Never Starts (selected candidates who failed to start for exclusive and non-exclusive )", "Fill rate - % of jobs filled vs % jobs received (exclusive) Note: Cancelled jobs excluded", "Fill Rate - % of jobs filled vs % jobs received(non-exclusive)", "Time to accept /to position fulfillment (exclusive)", "Time to accept /to position fulfillment (non-exclusive)", "%  of contractors that are meeting acceptable performance level.", "% of failed hires or ended prematurely due to negative reasons", "% of contractors fulfilling the anticipated duration of the job assignments (before the 3 year limit)", "Number of contractors presented who are identified as fraudulent", "No of resources leaving the account in the quarter/Headcount of the account at the end of the quarter"));
			ArrayList<Object> kpiList = getKPIList();

			ArrayList<String> greenVals = new ArrayList<>(Arrays.asList("<=10 days", "<=10 days", ">50%", ">50%", "<5%", "<5%", ">85%", "25%", "20 Business Days", "20 Business Days", ">95%", "<5%", ">90%", "0", "<=15%"));
			ArrayList<String> yellowVals = new ArrayList<>(Arrays.asList("N/A", "N/A", "<50-33%", "<50-33%", "5-10%", "5-10%", "<85-80%", "<24-20%", ">21 Business Days", ">21 Business Days", "<95-90%", "N/A", "<90-80%", " ", "N/A"));
			ArrayList<String> redVals = new ArrayList<>(Arrays.asList(">10 days", ">10 days", "<33%", "<33%", ">10%", ">10%", "<80%", "<20%", ">31 Business Days", ">31 Business Days", "<90%", ">5%", "<80%", ">0", ">15%"));
			int skip = 0;
			for (int row = 3; row < 18; row++) {
				Row rowObj = sheet.getRow(row);
				if (rowObj == null) {
					rowObj = sheet.createRow(row);
				}
				Cell cell = rowObj.createCell(1);
				cell.setCellStyle(ExcelStyleUtil.createBrownStyle(workbook));

				Cell cell2 = rowObj.createCell(3);
				cell2.setCellStyle(ExcelStyleUtil.createSLAheaders(workbook));
				if(row == 3) {
					cell2.setCellValue(kpiHeaders.get(0));
				}
				else if(row == 9) {
					cell2.setCellValue(kpiHeaders.get(1));
				}
				else if(row > 12 && row < 18) {
					cell2.setCellValue(kpiHeaders.get(2 + skip));
					skip++;
				}

				Cell cell3 = rowObj.createCell(4);
				cell3.setCellStyle(ExcelStyleUtil.createSLAdescriptors(workbook));
				cell3.setCellValue(kpiDescriptions.get(row - 3));
				sheet.setColumnWidth(4, 256 * 60);

				Cell cell4 = rowObj.createCell(5);
				Double kpiValue = extractDoubleFromString(kpiList.get(row - 3).toString());
				cell4.setCellValue(kpiList.get(row - 3).toString());
				setCellStyleBasedOnString(workbook, cell4, kpiValue, row);
				
				 // set Vendor Comments
				Cell cell13 = rowObj.createCell(6);
				if(row == 3) {
					cell13.setCellValue(Respond_With_Resume_Report.excCount + " Exclusive in Q" + quarter);
				} else if (row == 4) {
					cell13.setCellValue(Respond_With_Resume_Report.nonCount + " Non-Exclusive in Q" + quarter);
				} else if (row == 5) {
					cell13.setCellValue(Resumes_Interview_Exc_Report.countExc + " Exclusive in Q" + quarter);
				} else if (row == 6) {
					cell13.setCellValue(Resumes_Interview_NonExc_Report.countNon + " Non-Exclusive in Q" + quarter);
				} else if (row == 7) {
					cell13.setCellValue(Offers_Declined_Report.numOffers + " Offers, " + Offers_Declined_Report.declinedOffers + " Declination");
				} else if (row == 8) {
					cell13.setCellValue(Never_Starts_Report.acceptances + " Acceptances, " + Never_Starts_Report.rescinded + " Rescinded");
				} else if (row == 9) {
					cell13.setCellValue(Fill_Rate_Report.jobsFilledExc + " of " + Fill_Rate_Report.jobsReceivedExc + " in Q" + quarter);
				} else if (row == 10) {
					cell13.setCellValue(Fill_Rate_Report.jobsFilledNon + " of " + Fill_Rate_Report.jobsReceivedNon + " in Q" + quarter);
				} else if (row == 11) {
					cell13.setCellValue((int) Time_Accept_Report.fillsExc + " fills/" + (int) Time_Accept_Report.daysExc + " days");
				} else if (row == 12) {
					cell13.setCellValue((int) Time_Accept_Report.fillsNon + " fills/" + (int) Time_Accept_Report.daysNon + " days");
				} else if (row == 13) {
					cell13.setCellValue(Failed_Hires_Report.successfulHires + " of " + Failed_Hires_Report.headCount);
				} else if (row == 14) {
					cell13.setCellValue(Failed_Hires_Report.failedHires + " of " + Failed_Hires_Report.headCount);
				} else if (row == 15) {
					cell13.setCellValue(Completion_Rate_Report.numResources + " of " + Completion_Rate_Report.headCount);
				} else if (row == 16) {
					cell13.setCellValue(Resume_Fraud_Report.fraudCount);
				} else if (row == 17) {
					cell13.setCellValue(Workplace_Turbulence_Report.numResources + " of " + Workplace_Turbulence_Report.headCount);
				}
				cell13.setCellStyle(ExcelStyleUtil.sideSLAdescriptors(workbook));

				// set target column
				Cell cell6 = rowObj.createCell(13);
				cell6.setCellValue("Green");
				cell6.setCellStyle(ExcelStyleUtil.greenBackground(workbook));

				// set green column
				Cell cell7 = rowObj.createCell(14);
				cell7.setCellValue(greenVals.get(row - 3));
				cell7.setCellStyle(ExcelStyleUtil.greenBackground(workbook));

				// set yellow column
				Cell cell8 = rowObj.createCell(15);
				cell8.setCellValue(yellowVals.get(row - 3));
				cell8.setCellStyle(ExcelStyleUtil.yellowBackground(workbook));

				// set red column
				Cell cell9 = rowObj.createCell(16);
				cell9.setCellValue(redVals.get(row - 3));
				cell9.setCellStyle(ExcelStyleUtil.redBackground(workbook));

				// scorecard data collection columns
				Cell cell10 = rowObj.createCell(18);
				cell10.setCellValue("Supplier");
				cell10.setCellStyle(ExcelStyleUtil.sideSLAdescriptors(workbook));
				Cell cell11 = rowObj.createCell(19);
				cell11.setCellValue("Business will validate");
				cell11.setCellStyle(ExcelStyleUtil.sideSLAdescriptors(workbook));
				Cell cell12 = rowObj.createCell(20);
				cell12.setCellValue("Quarterly");
				cell12.setCellStyle(ExcelStyleUtil.sideSLAdescriptors(workbook));
			}
			sheet.setColumnWidth(6, 256 * 30);
			
			int indexVal = 0;
			// Loop through the quarters
			for(int quarterNum : quarters) {
				Main.KPIforQuarter(workbook, allSheets, quarterNum, newYears[indexVal], fileReqPath, fileEntechPath);
				ArrayList<Object> newKPIlist = getKPIList();
				for (int row = 3; row < 18; row++) {
					Row rowQuart = sheet.getRow(row);
					if (rowQuart == null) {
						rowQuart = sheet.createRow(row);
					}
					
					Double newKPIValue = extractDoubleFromString(newKPIlist.get(row - 3).toString());
					Cell cell5 = rowQuart.createCell(9 + indexVal);
					cell5.setCellValue(newKPIlist.get(row - 3).toString());
					setCellStyleBasedOnString(workbook, cell5, newKPIValue, row);
				}
				indexVal++;
			} 

			// Save input details for side bar 
			ArrayList<String> sideHeaders = new ArrayList<>(Arrays.asList("Who provides \ninformation?", "Data Source", "Frequency data provided to SPM"));

			for (int i = 18; i < 21; i++) {
				Cell c = r0.createCell(i);
				Cell h = r1.createCell(i);
				h.setCellValue(sideHeaders.get(i - 18));
				h.setCellStyle(ExcelStyleUtil.createBlueStyle(workbook));
				c.setCellValue("Scorecard Data Collection");
				c.setCellStyle(ExcelStyleUtil.createBorderGrayStyle(workbook));
				sheet.setColumnWidth(i, 256 * 15);
			}

			Cell c = r1.createCell(21);
			c.setCellStyle(ExcelStyleUtil.createBorderGrayStyle(workbook));		

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Double extractDoubleFromString(String str) {
		// Use regular expression to find a numeric value in the string
		Pattern pattern = Pattern.compile("-?\\d+(\\.\\d+)?");
		Matcher matcher = pattern.matcher(str);

		// If a numeric value is found, parse it to a double and return it
		if (matcher.find()) {
			return Double.parseDouble(matcher.group());
		}

		// If no numeric value is found, return 0.0 as a default
		return 0.0;
	}

	static class Threshold {
		double minValue;
		double maxValue;
		String style;

		Threshold(double minValue, double maxValue, String style) {
			this.minValue = minValue;
			this.maxValue = maxValue;
			this.style = style;
		}
	}

	public static void setCellStyleBasedOnString(XSSFWorkbook workbook, Cell cell, Double kpiValue, int rowNum) {
		// Define thresholds and styles for each row range
		Map<Integer, Threshold[]> thresholds = new HashMap<>();

		thresholds.put(3, new Threshold[]{new Threshold(0, 10, "green"), new Threshold(10, Double.MAX_VALUE, "red")});
		thresholds.put(4, new Threshold[]{new Threshold(0, 10, "green"), new Threshold(10, Double.MAX_VALUE, "red")});
		thresholds.put(5, new Threshold[]{new Threshold(50, Double.MAX_VALUE, "green"), new Threshold(33, 50, "yellow"), new Threshold(0, 33, "red")});
		thresholds.put(6, new Threshold[]{new Threshold(50, Double.MAX_VALUE, "green"), new Threshold(33, 50, "yellow"), new Threshold(0, 33, "red")});
		thresholds.put(7, new Threshold[]{new Threshold(0, 5, "green"), new Threshold(5, 10, "yellow"), new Threshold(10, Double.MAX_VALUE, "red")});
		thresholds.put(8, new Threshold[]{new Threshold(0, 5, "green"), new Threshold(5, 10, "yellow"), new Threshold(10, Double.MAX_VALUE, "red")});
		thresholds.put(9, new Threshold[]{new Threshold(85, Double.MAX_VALUE, "green"), new Threshold(80, 85, "yellow"), new Threshold(0, 80, "red")});
		thresholds.put(10, new Threshold[]{new Threshold(25, Double.MAX_VALUE, "green"), new Threshold(20, 24, "yellow"), new Threshold(0, 20, "red")});
		thresholds.put(11, new Threshold[]{new Threshold(0, 20, "green"), new Threshold(21, 31, "yellow"), new Threshold(31, Double.MAX_VALUE, "red")});
		thresholds.put(12, new Threshold[]{new Threshold(0, 20, "green"), new Threshold(21, 31, "yellow"), new Threshold(31, Double.MAX_VALUE, "red")});
		thresholds.put(13, new Threshold[]{new Threshold(95, Double.MAX_VALUE, "green"), new Threshold(90, 95, "yellow"), new Threshold(0, 90, "red")});
		thresholds.put(14, new Threshold[]{new Threshold(0, 5, "green"), new Threshold(5, Double.MAX_VALUE, "red")});
		thresholds.put(15, new Threshold[]{new Threshold(90, Double.MAX_VALUE, "green"), new Threshold(80, 90, "yellow"), new Threshold(0, 80, "red")});
		thresholds.put(16, new Threshold[]{new Threshold(0, 0, "green"), new Threshold(0, Double.MAX_VALUE, "red")});
		thresholds.put(17, new Threshold[]{new Threshold(0, 15, "green"), new Threshold(15, Double.MAX_VALUE, "red")});

		// Get thresholds for the current row
		Threshold[] rowThresholds = thresholds.get(rowNum);

		// Apply the appropriate style based on the KPI value
		for (Threshold threshold : rowThresholds) {
			if (kpiValue >= threshold.minValue && kpiValue <= threshold.maxValue) {
				switch (threshold.style) {
				case "green":
					cell.setCellStyle(ExcelStyleUtil.greenBackground(workbook));
					break;
				case "yellow":
					cell.setCellStyle(ExcelStyleUtil.yellowBackground(workbook));
					break;
				case "red":
					cell.setCellStyle(ExcelStyleUtil.redBackground(workbook));
					break;
				}
				break;
			}
		}
	}


	public ArrayList<Object> getKPIList() {
		ArrayList<Object> kpiList = new ArrayList<>(Arrays.asList(
				Respond_With_Resume_Report.KPIexc, 
				Respond_With_Resume_Report.KPInon, 
				Resumes_Interview_Exc_Report.KPI, 
				Resumes_Interview_NonExc_Report.KPI, 
				Offers_Declined_Report.KPI, 
				Never_Starts_Report.KPI, 
				Fill_Rate_Report.KPIexc, 
				Fill_Rate_Report.KPInon, 
				Time_Accept_Report.KPIexc, 
				Time_Accept_Report.KPInon, 
				Failed_Hires_Report.KPIaccepted, 
				Failed_Hires_Report.KPI, 
				Completion_Rate_Report.KPI, 
				Resume_Fraud_Report.KPI, 
				Workplace_Turbulence_Report.KPI
				));
		return kpiList;
	}
}
