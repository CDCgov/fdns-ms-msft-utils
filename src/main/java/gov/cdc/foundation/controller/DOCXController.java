package gov.cdc.foundation.controller;

import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import org.apache.log4j.Logger;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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

import gov.cdc.foundation.helper.LoggerHelper;
import gov.cdc.foundation.helper.MessageHelper;
import gov.cdc.helper.ErrorHandler;
import gov.cdc.helper.common.ServiceException;
import io.swagger.annotations.ApiOperation;
import io.swagger.annotations.ApiParam;

@Controller
@EnableAutoConfiguration
@RequestMapping("/api/1.0/docx")
public class DOCXController {

	private static final Logger logger = Logger.getLogger(DOCXController.class);

	@RequestMapping(
		value = "extract",
		method = RequestMethod.POST,
		produces = MediaType.TEXT_PLAIN_VALUE
	)
	@ApiOperation(
		value = "Extract text from DOCX",
		notes = "Extract text from DOCX"
	)
	@ResponseBody
	public ResponseEntity<?> extractDataToJson(
		@ApiParam(value = "DOCX File") @RequestParam("file") MultipartFile file,
		@ApiParam(value = "Expected file name") @RequestParam(value = "filename", required = false) String filename
	) {		
		Map<String, Object> log = new HashMap<String, Object>();
		log.put(MessageHelper.CONST_METHOD, MessageHelper.METHOD_EXTRACTDATA_DOCX);
		log.put(MessageHelper.CONST_FILENAME, file.getOriginalFilename());

		try {
			if (!file.getOriginalFilename().toLowerCase().endsWith(".docx"))
				throw new ServiceException("Only *.docx files are supported.");

			XWPFDocument doc = new XWPFDocument(file.getInputStream());
			XWPFWordExtractor extractor = new XWPFWordExtractor(doc);

			// Get filename
			String fn = filename == null || filename.isEmpty() ? UUID.randomUUID().toString() + ".txt" : filename;
			String headerValue = filename != null &&  !filename.isEmpty() ? "attachment; " : "";
			headerValue += "filename=" + fn;

			return ResponseEntity.status(HttpStatus.OK).header("Content-Disposition", headerValue).body(extractor.getText());
		} catch (Exception e) {
			logger.error(e);
			LoggerHelper.log(MessageHelper.METHOD_EXTRACTDATA_DOCX, log);
			
			return ErrorHandler.getInstance().handle(e, log);
		}
	}

}