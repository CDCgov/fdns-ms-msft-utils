package gov.cdc.foundation.helper;

import java.util.HashMap;
import java.util.Map;

import gov.cdc.helper.AbstractMessageHelper;

public class MessageHelper extends AbstractMessageHelper {

	public static final String CONST_FILENAME = "filename";

	public static final String METHOD_INDEX = "index";
	public static final String METHOD_GETSHEETS = "getSheets";
	public static final String METHOD_EXTRACTDATA_XLSX = "extractDataFromXLSX";
	public static final String METHOD_EXTRACTDATA_DOCX = "extractDataFromDOCX";
	public static final String METHOD_CONVERTCSVTOXLSX = "convertCSVToXLSX";

	private MessageHelper() {
		throw new IllegalAccessError("Helper class");
	}

	public static Map<String, Object> initializeLog(String method) {
		Map<String, Object> log = new HashMap<>();
		log.put(MessageHelper.CONST_METHOD, method);
		return log;
	}

}
