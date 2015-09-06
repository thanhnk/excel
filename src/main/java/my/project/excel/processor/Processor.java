package my.project.excel.processor;

public interface Processor {
	public void init(String inputFile, String outputFile, boolean isHugeFile)
			throws Exception;

	public void proccess() throws Exception;
}
