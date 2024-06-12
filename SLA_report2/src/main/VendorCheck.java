package main;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class VendorCheck {
	// Check if the vendor is valid
	public static boolean isValidVendor(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String value = cell.getStringCellValue();
			return "All".equals(value) ||
					"Cap/Entech".equals(value) ||
					"Cap/JTS/Entech".equals(value) ||
					"Entech".equals(value) ||
					"JTS/Entech".equals(value) ||
					"JTS/RT/Entech".equals(value);
		}
		return false;
	}

	// Check if the vendor is just Entech
	public static boolean isJustVendor(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String value = cell.getStringCellValue();
			return "Cap/Entech".equals(value) ||
					"Cap/JTS/Entech".equals(value) ||
					"Entech".equals(value) ||
					"JTS/Entech".equals(value) ||
					"JTS/RT/Entech".equals(value);
		}
		return false;
	}

	// Check if the vendor is "All"
	public static boolean isAllVendor(Cell cell) {
		if (cell != null && cell.getCellType() == CellType.STRING) {
			String value = cell.getStringCellValue();
			return "All".equals(value);
		}
		return false;
	}
}
