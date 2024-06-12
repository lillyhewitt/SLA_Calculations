package main;

import java.io.File;
import java.util.Scanner;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// calls classes to read the Vendor Req Report and write the SLA report 
public class Main {
	public static void main(String[] args) {
		// Increase the minimum inflate ratio
		ZipSecureFile.setMinInflateRatio(0.005);

		// ask user to input which quarter to generate report on
		Scanner scanner = new Scanner(System.in);
		int quarter = 0;
		// ask user for quarter, if user enters invalid quarter, ask again
		while (quarter == 0) {
			System.out.print("Please enter the quarter using just the number (ex. 1, 2, 3, 4) and click enter:\t");
			quarter = scanner.nextInt();

			if (quarter < 1 || quarter > 4) {
				System.out.println("Invalid quarter number. Please enter a number between 1 and 4.\n");
				quarter = 0;
			}
		}

		System.out.println();

		int year = 0;
		// ask user for year, if invalid year, ask again
		while (year == 0) {
			System.out.print("Please enter the year using format 20XX (ex. 2024) and click enter:\t");
			year = scanner.nextInt();
			scanner.nextLine(); // consume the newline character
			if (year < 2000) {
				System.out.println("Invalid year. Please enter a year after 2000.\n");
				year = 0;
			}
		}

		System.out.println();

		// ask user for files to read from
		System.out.println("Please enter the file path for the req report (ex. C:\\Users\\lhewitt\\Desktop\\file.xlsx):\t");
		String fileReqPath = scanner.nextLine();
		System.out.println("\nPlease enter the file path for Entech's own record-keeping (ex. C:\\Users\\lhewitt\\Desktop\\file.xlsx):\t");
		String fileEntechPath = scanner.nextLine(); 

		// Create file to save workbook to
		System.out.println("\nPlease enter full path to save new Excel Workbook to (ex. C:\\Users\\lhewitt\\Downloads\\Entech IT Staff MON SLA XQ20XX.xlsx) and click enter:\t");
		String filePathFinal = scanner.nextLine();
		File writeFile = new File(filePathFinal);  

		// Create Workbook
		XSSFWorkbook workbook = new XSSFWorkbook();

		// Create hashMap to store all sheets
		HashMap<XSSFSheet, String> allSheets = new HashMap<>();

		// set up file output stream and set to null
		FileOutputStream fos = null;

		try {
			createSheetsAndAddToMap(workbook, allSheets);

			// Run each sheet
			KPIforQuarter(workbook, allSheets, quarter, year, fileReqPath, fileEntechPath);

			// run SLA sheet
			XSSFSheet slaSheet = workbook.getSheet("SLA - Program Level");
			if (slaSheet != null) {
				SLA_Report report = new SLA_Report();
				report.write_SLA_Report(workbook, slaSheet, allSheets, quarter, year, fileReqPath, fileEntechPath);
			}

			// Run each sheet
			runQuarter(workbook, allSheets, quarter, year, fileReqPath, fileEntechPath);

			fos = new FileOutputStream(writeFile);
			workbook.write(fos);
			System.out.println("\nFile created and written");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
				if (workbook != null) {
					workbook.close();
				}
				scanner.close();
			} catch (IOException ex) {
				ex.printStackTrace();
			}
		}
	}

	public static void createSheetsAndAddToMap(XSSFWorkbook workbook, HashMap<XSSFSheet, String> allSheets) {
		XSSFSheet SLA_report = workbook.createSheet("SLA - Program Level");
		XSSFSheet fraudulent_resources = workbook.createSheet("Fraudulent Resources");
		XSSFSheet resume_to_interviews_exc = workbook.createSheet("Ratio Resumes to Interviews Exc");
		XSSFSheet resume_to_interviews_non = workbook.createSheet("Ratio Resumes to Interviews Non");
		XSSFSheet offers_declined = workbook.createSheet("# of Offers Declined");
		XSSFSheet fill_rate = workbook.createSheet("Fill Rate");
		XSSFSheet time_to_accept = workbook.createSheet("Time to Accept");
		XSSFSheet failed_hires = workbook.createSheet("Failed Hires");
		XSSFSheet completion_rate = workbook.createSheet("Assignment Completion Rate");
		XSSFSheet resume_fraud = workbook.createSheet("Resume Fraud");
		XSSFSheet workforce_turbulence = workbook.createSheet("Provider Workforce Turbulence");

		allSheets.put(SLA_report, "SLA - Program Level");
		allSheets.put(fraudulent_resources, "Fraudulent Resources");
		allSheets.put(resume_to_interviews_exc, "Ratio Resumes to Interviews Exc");
		allSheets.put(resume_to_interviews_non, "Ratio Resumes to Interviews Non");
		allSheets.put(offers_declined, "# of Offers Declined");
		allSheets.put(fill_rate, "Fill Rate");
		allSheets.put(time_to_accept, "Time to Accept");
		allSheets.put(failed_hires, "Failed Hires");
		allSheets.put(completion_rate, "Assignment Completion Rate");
		allSheets.put(resume_fraud, "Resume Fraud");
		allSheets.put(workforce_turbulence, "Provider Workforce Turbulence");
	}

	public static void KPIforQuarter(XSSFWorkbook workbook, HashMap<XSSFSheet, String> allSheets, int quarter, int year, String fileReqPath, String fileEntechPath) {
		// run Never_Starts_Report 
		Never_Starts_Report.read_From_Excel(fileEntechPath, quarter, year);
		// run Respond_With_Resume_Report
		Respond_With_Resume_Report.read_From_Excel(fileReqPath, quarter, year);

		// write each sheet in the workbookfor (Entry<XSSFSheet, String> entry : allSheets.entrySet()) {
		for (Entry<XSSFSheet, String> entry : allSheets.entrySet()) {
			String sheetName = entry.getValue();
			XSSFSheet sheet = entry.getKey();
			switch (sheetName) {
			case "Ratio Resumes to Interviews Exc":
				Resumes_Interview_Exc_Report.calculateKPI(quarter, year, fileReqPath);
				break;
			case "Ratio Resumes to Interviews Non":
				Resumes_Interview_NonExc_Report.calculateKPI(quarter, year, fileReqPath);
				break;
			case "# of Offers Declined":
				Offers_Declined_Report.calculateKPI(quarter, year, fileReqPath);
				break;
			case "Fill Rate":
				Fill_Rate_Report.calculateExcKPI(quarter, year, fileReqPath);
				Fill_Rate_Report.calculateNonKPI(quarter, year, fileReqPath);
				break;
			case "Time to Accept":
				Time_Accept_Report.calculateExcKPI(quarter, year, fileReqPath);
				Time_Accept_Report.calculateNonKPI(quarter, year, fileReqPath);
				break;
			case "Failed Hires":
				Failed_Hires_Report.calculateAcceptableKPI(quarter, year, fileEntechPath);
				Failed_Hires_Report.calculateKPI(quarter, year, fileEntechPath);
				break;
			case "Assignment Completion Rate":
				Completion_Rate_Report.calculateKPI(quarter, year, fileEntechPath);
				break;
			case "Resume Fraud":
				Resume_Fraud_Report.calculateKPI(quarter, year, fileEntechPath);
				break;
			case "Provider Workforce Turbulence":
				Workplace_Turbulence_Report.calculateKPI(quarter, year, fileEntechPath);
				break;
			}
		}
	}

	public static void runQuarter(XSSFWorkbook workbook, HashMap<XSSFSheet, String> allSheets, int quarter, int year, String fileReqPath, String fileEntechPath) {
		KPIforQuarter(workbook, allSheets, quarter, year, fileReqPath, fileEntechPath);
		// write each sheet in the workbook
		for (Entry<XSSFSheet, String> entry : allSheets.entrySet()) {
			String sheetName = entry.getValue();
			XSSFSheet sheet = entry.getKey();
			switch (sheetName) {
			case "Fraudulent Resources":
				Fraudulent_Report fraudReport = new Fraudulent_Report();
				fraudReport.write_Fraudulent_Resources(workbook, sheet, quarter, year);
				break;
			case "Ratio Resumes to Interviews Exc":
				Resumes_Interview_Exc_Report resumeIntExcReport = new Resumes_Interview_Exc_Report();
				resumeIntExcReport.write_Resumes_Interviews_Exc(workbook, sheet, quarter, year, fileReqPath);
				break;
			case "Ratio Resumes to Interviews Non":
				Resumes_Interview_NonExc_Report resumeIntNonExcReport = new Resumes_Interview_NonExc_Report();
				resumeIntNonExcReport.write_Resumes_Interviews_NonExc(workbook, sheet, quarter, year, fileReqPath);
				break;
			case "# of Offers Declined":
				Offers_Declined_Report offersDeclinedReport = new Offers_Declined_Report();
				offersDeclinedReport.write_Offers_Declined(workbook, sheet, quarter, year, fileReqPath);
				break;
			case "Fill Rate":
				Fill_Rate_Report fillRateReport = new Fill_Rate_Report();
				fillRateReport.write_Fill_Rate(workbook, sheet, quarter, year, fileReqPath);
				break;
			case "Time to Accept":
				Time_Accept_Report timeAcceptReport = new Time_Accept_Report();
				timeAcceptReport.write_Time_to_Accept(workbook, sheet, quarter, year, fileReqPath);
				break;
			case "Failed Hires":
				Failed_Hires_Report failedHiresReport = new Failed_Hires_Report();
				failedHiresReport.write_Failed_Hires(workbook, sheet, quarter, year, fileEntechPath);
				break;
			case "Assignment Completion Rate":
				Completion_Rate_Report completionRateReport = new Completion_Rate_Report();
				completionRateReport.write_Completion_Rate(workbook, sheet, quarter, year, fileEntechPath);
				break;
			case "Resume Fraud":
				Resume_Fraud_Report resumeFraudReport = new Resume_Fraud_Report();
				resumeFraudReport.write_Resume_Fraud(workbook, sheet, quarter, year, fileEntechPath);
				break;
			case "Provider Workforce Turbulence":
				Workplace_Turbulence_Report turbulenceReport = new Workplace_Turbulence_Report();
				turbulenceReport.write_Workplace_Turbulence(workbook, sheet, quarter, year, fileEntechPath);
				break;
			}
		}
	}
}
