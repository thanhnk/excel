package my.project.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		try {
			Date startTime = new Date();
			System.out.println("Started");
			File inputFile = new File(Constants.FILE_INPUT_PATH);
			File ouputFile = new File(Constants.FILE_OUPUT_PATH);

			Workbook workbook = new XSSFWorkbook(inputFile);
			Sheet spreadsheet = workbook.getSheet(Constants.SHEET_MAIN_NAME);

			// keep 100 rows in memory, exceeding rows will be flushed to disk;
			Workbook outputWorkbook = new SXSSFWorkbook(100);

			Sheet outputSheet = outputWorkbook
					.createSheet(Constants.SHEET_OUTPUT_NAME);

			Iterator<Row> rows = spreadsheet.iterator();
			int rowCount = 0;
			Row row = null;
			Row outputRow = null;
			Cell cell = null;
			Cell ouputCell = null;
			while (rows.hasNext()) {
				row = rows.next();
				outputRow = outputSheet.createRow(rowCount++);

				Iterator<Cell> cells = row.cellIterator();
				int cellCount = 0;
				while (cells.hasNext()) {
					cell = cells.next();
					ouputCell = outputRow.createCell(cellCount++);
					ouputCell.setCellValue(getCellStringValue(cell));
				}
			}
			outputWorkbook.write(new FileOutputStream(ouputFile));
			((SXSSFWorkbook) outputWorkbook).dispose();
			outputWorkbook.close();
			workbook.close();

			Date endTime = new Date();
			long diff = endTime.getTime() - startTime.getTime();
			long diffSeconds = diff / 1000;
			System.out.println("Period:" + diffSeconds + " (s)");

			System.out.println("Finish");
		} catch (IOException | InvalidFormatException e) {
			e.printStackTrace();
		}
	}

	private static String getCellStringValue(Cell cell) {
		String value = "";
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			value = cell.getNumericCellValue() + "";
			break;
		case Cell.CELL_TYPE_STRING:
			value = cell.getStringCellValue();
			break;
		}
		return value;
	}
}
