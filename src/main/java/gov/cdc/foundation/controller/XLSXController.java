package gov.cdc.foundation.controller;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.UUID;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVRecord;
import org.apache.commons.io.FilenameUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.fasterxml.jackson.databind.ObjectMapper;

import gov.cdc.foundation.helper.LoggerHelper;
import gov.cdc.foundation.helper.MessageHelper;
import gov.cdc.helper.ErrorHandler;
import gov.cdc.helper.common.ServiceException;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;
import kotlin.text.Regex;

@Controller
@EnableAutoConfiguration
@RequestMapping("/api/1.0/xlsx")
public class XLSXController {

	private static final Logger logger = Logger.getLogger(XLSXController.class);

	@RequestMapping(value = "sheets", method = RequestMethod.POST, produces = MediaType.APPLICATION_JSON_VALUE)
	@ApiOperation(value = "Get the list of sheets", notes = "Get the list of sheets of a XLSX file.")
	@ResponseBody
	public ResponseEntity<?> getSheets(@ApiParam(value = "XLSX File") @RequestParam("file") MultipartFile file) throws IOException {
		ObjectMapper mapper = new ObjectMapper();

		Map<String, Object> log = new HashMap<String, Object>();
		log.put(MessageHelper.CONST_METHOD, MessageHelper.METHOD_GETSHEETS);
		log.put(MessageHelper.CONST_FILENAME, file.getOriginalFilename());

		Workbook wb = null;

		try {
			if (!file.getOriginalFilename().toLowerCase().endsWith(".xlsx"))
				throw new ServiceException("Only *.xlsx files are supported.");

			// Build the JSON
			JSONArray arr = new JSONArray();
			wb = WorkbookFactory.create(file.getInputStream());
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				Sheet sheet = wb.getSheetAt(i);
				JSONObject obj = new JSONObject();
				obj.put("name", sheet.getSheetName());
				obj.put("index", i);
				arr.put(obj);
			}

			JSONObject result = new JSONObject();
			result.put("items", arr);
			result.put("total", arr.length());

			return ResponseEntity.status(HttpStatus.OK).body(mapper.readTree(result.toString()));
		} catch (Exception e) {
			logger.error(e);
			LoggerHelper.log(MessageHelper.METHOD_GETSHEETS, log);

			return ErrorHandler.getInstance().handle(e, log);
		} finally {
			if (wb != null)
				wb.close();
		}
	}

	@RequestMapping(value = "extract/json", method = RequestMethod.POST, produces = MediaType.APPLICATION_JSON_VALUE)
	@ApiOperation(value = "Extract data from XLSX to JSON", notes = "Extract data from XLSX to JSON")
	@ResponseBody
	public ResponseEntity<?> extractDataToJson(@ApiParam(value = "XLSX File") @RequestParam("file") MultipartFile file,
						  @ApiParam(value = "Sheet Name") @RequestParam(value = "sheetName", required = false) String sheetName,
						  @ApiParam(value = "Sheet Range like A1:D1 or A2:A10") @RequestParam(value = "sheetRange") String sheetRange,
						  @ApiParam(value = "Orientation", allowableValues = "portrait,landscape") @RequestParam(value = "orientation", required = false, defaultValue = "portrait") String orientation,
						  @ApiParam(value = "Expected file name") @RequestParam(value = "filename", required = false) String filename) throws IOException {

		ObjectMapper mapper = new ObjectMapper();

		Map<String, Object> log = new HashMap<String, Object>();
		log.put(MessageHelper.CONST_METHOD, MessageHelper.METHOD_EXTRACTDATA_XLSX);
		log.put(MessageHelper.CONST_FILENAME, file.getOriginalFilename());

		Workbook wb = null;

		try {
			if (!file.getOriginalFilename().toLowerCase().endsWith(".xlsx"))
				throw new ServiceException("Only *.xlsx files are supported.");

			wb = WorkbookFactory.create(file.getInputStream());

			// Get sheet
			Sheet s = sheetName == null || sheetName.isEmpty() ? wb.getSheetAt(0) : wb.getSheet(sheetName);

			if (s == null)
				throw new ServiceException("The following sheet doesn't exist: " + sheetName);

			// Get data
			JSONArray data = extractData(s, sheetRange == null ? "" : sheetRange, orientation == null || orientation.isEmpty() ? "portrait" : orientation);

			// Get filename
			String fn = filename == null || filename.isEmpty() ? UUID.randomUUID().toString() + ".json" : filename;
			String headerValue = filename != null && !filename.isEmpty() ? "attachment; " : "";
			headerValue += "filename=" + fn;

			JSONObject json = new JSONObject();
			json.put("rows", data.length());
			json.put("cols", data.length() > 0 ? data.getJSONArray(0).length() : 0);
			json.put("items", data);

			return ResponseEntity.status(HttpStatus.OK).header(
					"Content-Disposition", headerValue
			).body(mapper.readTree(json.toString()));
		} catch (Exception e) {
			logger.error(e);
			LoggerHelper.log(MessageHelper.METHOD_EXTRACTDATA_XLSX, log);

			return ErrorHandler.getInstance().handle(e, log);
		} finally {
			if (wb != null)
				wb.close();
		}
	}


	@RequestMapping(value = "extract/csv", method = RequestMethod.POST, produces = "text/csv")
	@ApiOperation(value = "Extract data from XLSX to CSV", notes = "Extract data from XLSX to CSV")
	@ResponseBody
	public ResponseEntity<?> extractDataToCsv(@ApiParam(value = "XLSX File") @RequestParam("file") MultipartFile file,
						 @ApiParam(value = "Sheet Name") @RequestParam(value = "sheetName", required = false) String sheetName,
						 @ApiParam(value = "Sheet Range like A1:D1 or A2:A10") @RequestParam(value = "sheetRange") String sheetRange,
						 @ApiParam(value = "Orientation", allowableValues = "portrait,landscape") @RequestParam(value = "orientation", required = false, defaultValue = "portrait") String orientation,
						 @ApiParam(value = "Expected file name") @RequestParam(value = "filename", required = false) String filename) throws IOException {

		Map<String, Object> log = new HashMap<String, Object>();
		log.put(MessageHelper.CONST_METHOD, MessageHelper.METHOD_EXTRACTDATA_XLSX);
		log.put(MessageHelper.CONST_FILENAME, file.getOriginalFilename());

		Workbook wb = null;

		try {
			if (!file.getOriginalFilename().toLowerCase().endsWith(".xlsx"))
				throw new ServiceException("Only *.xlsx files are supported.");

			wb = WorkbookFactory.create(file.getInputStream());

			// Get sheet
			Sheet s = sheetName == null || sheetName.isEmpty() ? wb.getSheetAt(0) : wb.getSheet(sheetName);

			if (s == null)
				throw new ServiceException("The following sheet doesn't exist: " + sheetName);

			// Get data
			JSONArray data = extractData(s, sheetRange == null ? "" : sheetRange, orientation == null || orientation.isEmpty() ? "portrait" : orientation);

			// Get csv
			StringBuilder csv = new StringBuilder();
			for (int i = 0; i < data.length(); i++) {
				JSONArray row = data.getJSONArray(i);
				int nbOfColums = row.length();
				for (int j = 0; j < nbOfColums; j++) {
					String cell = row.getString(j);
					csv.append("\"" + cell + "\"");
					if (j < nbOfColums - 1) csv.append(',');
				}
				csv.append('\n');
			}

			// Get filename
			String fn = filename == null || filename.isEmpty() ? UUID.randomUUID().toString() + ".csv" : filename;
			String headerValue = filename != null && !filename.isEmpty() ? "attachment; " : "";
			headerValue += "filename=" + fn;

			return ResponseEntity.status(HttpStatus.OK).header(
					"Content-Disposition", headerValue
			).body(csv);
		} catch (Exception e) {
			logger.error(e);
			LoggerHelper.log(MessageHelper.METHOD_EXTRACTDATA_XLSX, log);

			return ErrorHandler.getInstance().handle(e, log);
		} finally {
			if (wb != null)
				wb.close();
		}
	}

	@RequestMapping(value = "from/csv", method = RequestMethod.POST, produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	@ApiOperation(value = "Transform a CSV to XLSX", notes = "Transform a CSV to XLSX")
	@ResponseBody
	public ResponseEntity<?> convertCSVToXLSX(@ApiParam(value = "CSV File") @RequestParam("file") MultipartFile file,
						 @ApiParam(value = "Expected file name") @RequestParam(value = "filename", required = false) String filename) throws IOException {

		Map<String, Object> log = new HashMap<String, Object>();
		log.put(MessageHelper.CONST_METHOD, MessageHelper.METHOD_CONVERTCSVTOXLSX);
		log.put(MessageHelper.CONST_FILENAME, file.getOriginalFilename());

		Workbook wb = null;

		try {
			if (!file.getOriginalFilename().toLowerCase().endsWith(".csv"))
				throw new ServiceException("Only *.csv files are supported.");

			wb = new SXSSFWorkbook();
			Sheet s = wb.createSheet(FilenameUtils.getBaseName(file.getOriginalFilename()));

			int rowNum = 0;
			BufferedReader br = new BufferedReader(new InputStreamReader(file.getInputStream()));
			Iterable<CSVRecord> records = CSVFormat.EXCEL.parse(br);
			for (CSVRecord record : records) {
				Row currentRow = s.createRow(rowNum);
				for( int i = 0; i < record.size(); i ++ ) {
					currentRow.createCell(i).setCellValue(record.get(i));
				}
				rowNum++;
			}

			// Get filename
			String fn = filename == null || filename.isEmpty() ? FilenameUtils.getBaseName(file.getOriginalFilename()) + ".xlsx" : filename;
			String headerValue = filename != null && !filename.isEmpty() ? "attachment; " : "";
			headerValue += "filename=" + fn;

			ByteArrayOutputStream bos = new ByteArrayOutputStream();
			try {
				wb.write(bos);
				return ResponseEntity.status(HttpStatus.OK).header(
						"Content-Disposition", headerValue
						).body(bos.toByteArray());
			} finally {
				bos.close();
				wb.close();
			}

		} catch (Exception e) {
			logger.error(e);
			LoggerHelper.log(MessageHelper.METHOD_CONVERTCSVTOXLSX, log);

			return ErrorHandler.getInstance().handle(e, log);
		} finally {
			if (wb != null)
				wb.close();
		}
	}

	private JSONArray extractData(Sheet s, String range, String orientation) throws ServiceException {
		// Check the range syntax
		Regex regex = new Regex("[A-Z]+\\d+:[A-Z]+\\d+");
		if (regex.matchEntire(range) == null)
			throw new ServiceException("The sheet range expression is not valid.");

		// Get row and cols indexes
		int startCol = index(new Regex("[A-Z]+").find(range, 0).getGroupValues().get(0)) - 1;
		int endCol = index(new Regex("[A-Z]+").find(range, range.indexOf(":")).getGroupValues().get(0)) - 1;
		int startRow = Integer.parseInt(new Regex("\\d+").find(range, 0).getGroupValues().get(0)) - 1;
		int endRow = Integer.parseInt(new Regex("\\d+").find(range, range.indexOf(":")).getGroupValues().get(0)) - 1;

		if (endCol < startCol)
			throw new ServiceException("The end column needs to be after the start column.");
		if (endRow < startRow)
			throw new ServiceException("The end column needs to be after the start column.");
		if (orientation.toLowerCase().equals("portrait") && startRow != endRow)
			throw new ServiceException("If the mode `portrait` is selected, the start and end rows must be the same.");
		if (orientation.toLowerCase().equals("landscape") && startCol != endCol)
			throw new ServiceException("If the mode `landscape` is selected, the start and end columns must be the same.");

		JSONArray arr = new JSONArray();

		if (orientation.toLowerCase().equals("portrait")) {
			boolean c = true;
			int rowIdx = startRow;
			while (c) {
				Row r = s.getRow(rowIdx);
				if (r != null) {
					JSONArray row = new JSONArray();
					for (int colIdx = startCol; colIdx <= endCol; colIdx++)
						row.put(cellToStr(r.getCell(colIdx)));

					boolean emptyLine = true;
					for (int i = 0; i < row.length(); i++) {
						Object value = row.get(i);
						emptyLine = emptyLine && (value == null || (value instanceof String && ((String) value).isEmpty()));
					}
					c = !emptyLine;
					if (c)
						arr.put(row);
					rowIdx++;
				} else
					c = false;
			}
		} else {
			boolean c = true;

			// First, we need to look for the column where all rows are blank
			int colIdxMax = startCol;
			while (c) {
				Set<String> values = new HashSet<String>();
				for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
					Row r = s.getRow(rowIdx);
					if (r != null)
						values.add(cellToStr(r.getCell(colIdxMax)));
					else
						values.add("");
				}
				boolean emptyCol = true;
				for (String value : values)
					emptyCol = emptyCol && value.isEmpty();
				c = !emptyCol;
				if (c)
					colIdxMax++;
			}
			colIdxMax--;

			// Then extract
			for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
				Row r = s.getRow(rowIdx);
				JSONArray row = new JSONArray();
				for (int colIdx = startCol; colIdx <= colIdxMax; colIdx++) {
					if (r != null)
						row.put(cellToStr(r.getCell(colIdx)));
					else
						row.put("");
				}
				arr.put(row);
			}
		}
		return arr;
	}

	private int index(String letterIndex) {
		if (letterIndex == null || letterIndex.isEmpty())
			return 0;
		else {
			int index = 0;
			int position = 1;
			for (int i = 0; i < letterIndex.length(); i ++) {
				char letter = letterIndex.charAt(i);
				index += ((int) letter - (int)'A' + 1) * Math.pow(26.0, letterIndex.length() - (double)(position));
				position++;
			}
			return index;
		}
	}

	private String cellToStr(Cell cell) {
		if (cell == null)
			return "";
		else {
			switch (cell.getCellTypeEnum()) {
				case NUMERIC :
					if (HSSFDateUtil.isCellDateFormatted(cell))
						return dateCellToStr(cell);
					else
						return Double.toString(cell.getNumericCellValue());
				case BOOLEAN:
					return Boolean.toString(cell.getBooleanCellValue());
				case STRING:
					return cell.getStringCellValue();
				case FORMULA:
					switch (cell.getCachedFormulaResultTypeEnum()) {
						case NUMERIC :
							if (HSSFDateUtil.isCellDateFormatted(cell))
								return dateCellToStr(cell);
							else
								return Double.toString(cell.getNumericCellValue());
						case BOOLEAN:
							return Boolean.toString(cell.getBooleanCellValue());
						case STRING:
							return cell.getStringCellValue();
						default:
							return "";
					}
				default:
					return "";
			}
		}
	}

	private String dateCellToStr(Cell c) {
		Date date = HSSFDateUtil.getJavaDate(c.getNumericCellValue());
		String dateFmt = c.getCellStyle().getDataFormatString();
		return new CellDateFormatter(dateFmt).format(date);
	}


}