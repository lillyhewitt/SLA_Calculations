package main;

import java.util.ArrayList;
import java.util.Arrays;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Fraudulent_Report {

	// Write Fraudulent Resources to Excel file
	public void write_Fraudulent_Resources(XSSFWorkbook workbook, XSSFSheet sheet, int quarter, int year) {
		try {
			// Add title to SLA report
			Row r0 = sheet.createRow(0);
			r0.setHeightInPoints(50); // Set row height to accommodate wrapped text
			Cell c0 = r0.createCell(0);
			c0.setCellValue("Entech - Q1 SLAs \nResources Identified as Fraudulent");
			c0.setCellStyle(ExcelStyleUtil.createNavyBlueStyle(workbook));

			// Add merged regions for title
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 10));

			// Add Column headers
			Row r1 = sheet.createRow(3);
			ArrayList<String> columnHeaders = new ArrayList<>(Arrays.asList(" ", "Start Date", "End Date", "Subdivision", "Cost Center", "Bill Rate", "Hiring Manager", "Skill Set", "Additional Comments"));
			for (int col = 1; col < columnHeaders.size(); col++) { // Fix here, start from index 1
				Cell c = r1.createCell(col+1);
				c.setCellValue(columnHeaders.get(col));
				c.setCellStyle(ExcelStyleUtil.createFraudHeaderStyle(workbook));
				// Set column width to fit the content
				sheet.setColumnWidth(col+1, 6000); // Adjust width as necessary
			}

			// Create tables per quarter
			ArrayList<String> quarters = new ArrayList<>(Arrays.asList("Q1", "Q2", "Q3", "Q4"));
			for (int i = 0; i < quarters.size(); i++) { // Fix the loop index, starting from 0
				int startRow = i * 4 + 4;
				int endRow = startRow + 2;

				// Create rows and cells covering the merged region, and apply the border style
				for (int j = startRow; j <= endRow; j++) {
					Row row = sheet.createRow(j);
					for (int k = 1; k <= 9; k++) { // Skip column 0
						Cell cell = row.createCell(k);
						cell.setCellStyle(ExcelStyleUtil.createPlainTableStyle(workbook));
					}
				}

				// Merge cells
				sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 1, 1));

				// Set the value of the cell at the start of the merged region
				Row startRowObject = sheet.getRow(startRow);
				Cell cell = startRowObject.createCell(1); // Make sure to create cell at column 1
				cell.setCellValue(quarters.get(i));
				cell.setCellStyle(ExcelStyleUtil.createTableStyle(workbook));
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
