package my.project.excel;

import my.project.excel.model.Director;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class DataUtil {
	private DataUtil() {
	}

	public static void generateRow(String[] data, Sheet sheet, int rowOrder) {
		Row row = sheet.createRow(rowOrder);
		int cellOrder = 0;
		Cell cell = null;
		for (String item : data) {
			cell = row.createCell(cellOrder++);
			cell.setCellValue(item);
		}
	}

	public static Director map(Row row) {
		Director director = new Director();
		director.setDirectorId(getCellStringValue(row.getCell(0)));
		director.setCountry(getCellStringValue(row.getCell(1)));
		director.setCompanyId(getCellStringValue(row.getCell(2)));
		director.setInstitutionName(getCellStringValue(row.getCell(3)));
		director.setCompanyType(getCellStringValue(row.getCell(4)));
		director.setQualification(getCellStringValue(row.getCell(5)));
		director.setQualificationDes(getCellStringValue(row.getCell(6)));
		director.setQualificationDate(getCellStringValue(row.getCell(7)));
		director.setChecked(getCellStringValue(row.getCell(8)));
		return director;
	}

	public static String getCellStringValue(Cell cell) {
		if (cell == null) {
			return "";
		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		return cell.getStringCellValue();
		// String value = "";
		// switch (cell.getCellType()) {
		// case Cell.CELL_TYPE_NUMERIC:
		// value = cell.getNumericCellValue() + "";
		// break;
		// case Cell.CELL_TYPE_STRING:
		// value = cell.getStringCellValue();
		// break;
		// }
		// return value;
	}
}
