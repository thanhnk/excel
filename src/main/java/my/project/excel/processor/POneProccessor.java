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
	private String outputFile;
	private boolean inited = false;
	private boolean isHugeFile;
	private List<Director> directors;

	@Inject
	private Logger logger;

	private int limitedSimilarDirectors = 100;

	public final static String[] OUTPUT_HEADER = { "Director ID",
			"Connected ID", "Company ID" };

	@Override
	public void init(String inputFile, String outputFile, boolean isHugeFile)
			throws Exception {
		logger.info("Start init()");
		this.inputFile = inputFile;
		this.outputFile = outputFile;
		this.isHugeFile = isHugeFile;
		load();
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

		Workbook outputWorkbook = null;
		try {
			if (isHugeFile) {
				// keep 100 rows in memory, exceeding rows will be flushed to
				// disk;
				outputWorkbook = new SXSSFWorkbook(100);
				// outputWorkbook.setCompressTempFiles(true);
			} else {
				outputWorkbook = new XSSFWorkbook();
			}
			Sheet outputSheet = outputWorkbook
					.createSheet(Constants.SHEET_OUTPUT_NAME);
			generateHeader(outputSheet);
			int rowOrder = 1;
			for (Director director : directors) {
				processRow(director, outputSheet, rowOrder++);
			}
			outputSheet.setColumnWidth(0, 5000);
			outputSheet.setColumnWidth(2, 5000);
			outputSheet.setColumnWidth(1, 15000);

			outputWorkbook.write(new FileOutputStream(new File(outputFile)));
		} catch (Exception e) {
			throw e;
		} finally {
			if (outputWorkbook != null) {
				if (isHugeFile) {
					((SXSSFWorkbook) outputWorkbook).dispose();
				}
				outputWorkbook.close();
			}
			logger.info("End process");
		}
	}

	private void processRow(Director director, Sheet sheet, int rowOrder)
			throws Exception {
		List<String> similarDirectors = new ArrayList<String>();
		int countSimilarDirector = 0;
		for (Director otherDirector : directors) {
			if (isSameCompany(director, otherDirector)) {
				similarDirectors.add(otherDirector.getDirectorId());
				countSimilarDirector++;
				if (countSimilarDirector > limitedSimilarDirectors) {
					logger.warning("Too many directors with the same company: "
							+ String.join(", ", similarDirectors));
					break;
				}
			}
		}
		String[] ouputData = { director.getDirectorId(),
				String.join(", ", similarDirectors), director.getCompanyId() };

		DataUtil.generateRow(ouputData, sheet, rowOrder);
		logger.info("Generated row " + rowOrder + ":"
				+ String.join("|", ouputData));
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

	private void load() throws Exception {
		Workbook inputWorkbook = new XSSFWorkbook(new File(inputFile));
		Sheet inputSheet = inputWorkbook.getSheet(Constants.SHEET_MAIN_NAME);
		inputSheet.removeRow(inputSheet.getRow(0)); // remove header line
		directors = new ArrayList<Director>(inputSheet.getLastRowNum());
		Iterator<Row> rows = inputSheet.iterator();
		Director director = null;
		while (rows.hasNext()) {
			director = DataUtil.map(rows.next());
			directors.add(director);
			logger.info("Imported:" + director);
		}
		inputWorkbook.close();
	}
}
