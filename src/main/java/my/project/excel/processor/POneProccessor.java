package my.project.excel.processor;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Logger;

import javax.inject.Inject;

import my.project.excel.Constants;
import my.project.excel.DataUtil;
import my.project.excel.model.Director;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POneProccessor implements Processor {
	private String inputFile;
	private String outputFileName;
	private String outputFileExtension;
	private boolean inited = false;
	private boolean isHugeFile;
	private List<Director> directors;
	private int outputRowOrder;
	private Sheet currentSheet;
	private Workbook outputWorkbook;
	private String startRow;
	private String endRow;

	@Inject
	private Logger logger;

	public final static String[] OUTPUT_HEADER = { "Director ID",
			"Connected ID", "Company ID" };

	@Override
	public void init(String inputFile, String outputFile, boolean isHugeFile)
			throws Exception {
		logger.info("Start init()");
		this.inputFile = inputFile;
		this.outputFileName = outputFile.substring(0,
				outputFile.indexOf(".xlsx"));
		this.outputFileExtension = ".xlsx";
		this.isHugeFile = isHugeFile;
		loadData();
		inited = true;
		logger.info("Imported " + directors.size() + " directors");
		logger.info("End init()");
	}

	@Override
	public void proccess() throws Exception {
		logger.info("Start process");
		if (!inited) {
			throw new Exception("The proccessor.init() need to call first");
		}

		try {
			generateNewOutput();
			int directorOrder = 1;
			for (Director director : directors) {
				if (StringUtils.isEmpty(startRow)) {
					startRow = "" + directorOrder;
				}
				endRow = "" + directorOrder;
				processDirector(director);
				logger.info("Processed row " + directorOrder++);
			}
			saveToFile();
		} catch (Exception e) {
			throw e;
		}
		logger.info("End process");
	}

	private void saveToFile() throws Exception {

		try {
			String fileName = outputFileName + "_" + startRow + "_" + endRow
					+ outputFileExtension;
			outputWorkbook.write(new FileOutputStream(new File(fileName)));
			logger.info("Saved file:" + fileName);
		} catch (Exception e) {
			throw e;
		} finally {
			if (outputWorkbook != null) {
				if (isHugeFile) {
					((SXSSFWorkbook) outputWorkbook).dispose();
				}
				outputWorkbook.close();
				outputWorkbook = null;
			}
			logger.info("End process");
		}
	}

	private void generateNewOutput() throws Exception {
		if (isHugeFile) {
			// keep 100 rows in memory, exceeding rows will be flushed to
			// disk;
			outputWorkbook = new SXSSFWorkbook(100);
			// outputWorkbook.setCompressTempFiles(true);
		} else {
			outputWorkbook = new XSSFWorkbook();
		}
		currentSheet = outputWorkbook.createSheet(Constants.SHEET_OUTPUT_NAME);
		generateHeader(currentSheet);
		currentSheet.setColumnWidth(0, 5000);
		currentSheet.setColumnWidth(2, 5000);
		currentSheet.setColumnWidth(1, 5000);
		outputRowOrder = 1;
		startRow = "";
		endRow = "";
	}

	private void processDirector(Director director) throws Exception {
		for (Director otherDirector : directors) {
			if (isSameCompany(director, otherDirector)) {
				String[] ouputData = { director.getDirectorId(),
						otherDirector.getDirectorId(), director.getCompanyId() };
				DataUtil.generateRow(ouputData, currentSheet, outputRowOrder++);
				if (outputRowOrder >= Constants.MAX_ROW_NUMBER_EXCEL) {
					saveToFile();
					generateNewOutput();
				}
			}

		}
	}

	private boolean isSameCompany(Director director1, Director director2) {
		return StringUtils.isNoneEmpty(director1.getCompanyId())
				&& !director1.getDirectorId().equals(director2.getDirectorId())
				&& director1.getCompanyId().equals(director2.getCompanyId());
	}

	private void generateHeader(Sheet sheet) throws Exception {
		int headerRowOrder = 0;
		DataUtil.generateRow(OUTPUT_HEADER, sheet, headerRowOrder);
	}

	private void loadData() throws Exception {
		Workbook inputWorkbook = new XSSFWorkbook(new File(inputFile));
		Sheet inputSheet = inputWorkbook.getSheet(Constants.SHEET_MAIN_NAME);
		inputSheet.removeRow(inputSheet.getRow(0)); // remove header line
		directors = new ArrayList<Director>(inputSheet.getLastRowNum());
		Iterator<Row> rows = inputSheet.iterator();
		Director director = null;
		int rowOrder = 1;
		logger.info("Start loading file");
		while (rows.hasNext()) {
			director = DataUtil.map(rows.next());
			directors.add(director);
			logger.info("Imported row " + (rowOrder++) + ":" + director);
		}
		inputWorkbook.close();
	}
}
