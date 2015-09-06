package my.project.excel;

import java.util.Date;

import my.project.excel.processor.Processor;

import org.apache.poi.ss.usermodel.Cell;
import org.jboss.weld.environment.se.Weld;
import org.jboss.weld.environment.se.WeldContainer;

public class Main {

	public static void main(String[] args) {
		try {
			Date startTime = new Date();
			System.out.println("Started");

			Weld weld = new Weld();
			WeldContainer container = weld.initialize();

			Processor processor = container.instance().select(Processor.class)
					.get();
			processor.init(Constants.FILE_INPUT_PATH,
					Constants.FILE_OUPUT_PATH, true);
			processor.proccess();

			Date endTime = new Date();
			long diff = endTime.getTime() - startTime.getTime();
			long diffSeconds = diff / 1000;
			System.out.println("Period:" + diffSeconds + " (s)");

			System.out.println("Finish");
		} catch (Exception e) {
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
