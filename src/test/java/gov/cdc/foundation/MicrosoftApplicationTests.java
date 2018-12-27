package gov.cdc.foundation;

import static org.assertj.core.api.Assertions.assertThat;
import static org.junit.Assert.assertThat;
import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.io.IOUtils;
import org.hamcrest.CoreMatchers;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.context.SpringBootTest.WebEnvironment;
import org.springframework.boot.test.web.client.TestRestTemplate;
import org.springframework.http.ResponseEntity;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;
import org.springframework.test.web.servlet.request.MockMultipartHttpServletRequestBuilder;
import org.springframework.test.web.servlet.request.MockMvcRequestBuilders;
import org.springframework.test.web.servlet.result.MockMvcResultMatchers;

import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

@RunWith(SpringRunner.class)
@SpringBootTest(webEnvironment = WebEnvironment.RANDOM_PORT, properties = {
		"logging.fluentd.host=fluentd", 
		"logging.fluentd.port=24224", 
		"proxy.hostname=", 
		"security.oauth2.resource.user-info-uri=", 
		"security.oauth2.protected=",
		"security.oauth2.client.client-id=",
		"security.oauth2.client.client-secret=",
		"ssl.verifying.disable=false" })
@AutoConfigureMockMvc
public class MicrosoftApplicationTests {

	@Autowired
	private TestRestTemplate restTemplate;
	@Autowired
	private MockMvc mvc;
	private String baseUrlPath = "/api/1.0/";

	@Test
	public void indexPage() {
		ResponseEntity<String> response = this.restTemplate.getForEntity("/", String.class);
		assertThat(response.getStatusCodeValue()).isEqualTo(200);
		assertThat(response.getBody(), CoreMatchers.containsString("FDNS Microsoft Utilities Microservice"));
	}

	@Test
	public void indexAPI() {
		ResponseEntity<String> response = this.restTemplate.getForEntity(baseUrlPath, String.class);
		assertThat(response.getStatusCodeValue()).isEqualTo(200);
		assertThat(response.getBody(), CoreMatchers.containsString("version"));
	}

	@Test
	public void getSheets() throws Exception {
		MockMultipartFile file = new MockMultipartFile("file", "sample.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", getResourceAsByte("/junit/sample.xlsx"));
		MockMultipartHttpServletRequestBuilder builder = MockMvcRequestBuilders.fileUpload(baseUrlPath + "/xlsx/sheets");
		MvcResult result = mvc.perform(builder.file(file)).andExpect(MockMvcResultMatchers.status().isOk()).andReturn();
		JsonElement json = new JsonParser().parse(result.getResponse().getContentAsString());
		assertTrue(json.getAsJsonObject().get("total").getAsInt() == 1);
		assertTrue(json.getAsJsonObject().get("items").getAsJsonArray().get(0).getAsJsonObject().get("name").getAsString().equals("Sheet1"));
	}

	@Test
	public void extractXlsxToCsv() throws Exception {
		MockMultipartFile file = new MockMultipartFile("file", "sample.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", getResourceAsByte("/junit/sample.xlsx"));
		MockMultipartHttpServletRequestBuilder builder = MockMvcRequestBuilders.fileUpload(baseUrlPath + "/xlsx/extract/csv?sheetRange=A1:C1");
		MvcResult result = mvc.perform(builder.file(file)).andExpect(MockMvcResultMatchers.status().isOk()).andReturn();
		String[] lines = result.getResponse().getContentAsString().split("\n");
		assertTrue(lines[0].equals("\"A1\",\"B1\",\"C1\""));
		assertTrue(lines[1].equals("\"A2\",\"B2\",\"C2\""));
	}

	@Test
	public void extractXlsxToJson() throws Exception {
		MockMultipartFile file = new MockMultipartFile("file", "sample.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", getResourceAsByte("/junit/sample.xlsx"));
		MockMultipartHttpServletRequestBuilder builder = MockMvcRequestBuilders.fileUpload(baseUrlPath + "/xlsx/extract/json?sheetRange=A1:C1");
		MvcResult result = mvc.perform(builder.file(file)).andExpect(MockMvcResultMatchers.status().isOk()).andReturn();
		JsonElement json = new JsonParser().parse(result.getResponse().getContentAsString());
		assertTrue(json.getAsJsonObject().get("rows").getAsInt() == 2);
		assertTrue(json.getAsJsonObject().get("cols").getAsInt() == 3);
	}

	@Test
	public void extractDocToTxt() throws Exception {
		MockMultipartFile file = new MockMultipartFile("file", "sample.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", getResourceAsByte("/junit/sample.docx"));
		MockMultipartHttpServletRequestBuilder builder = MockMvcRequestBuilders.fileUpload(baseUrlPath + "/docx/extract");
		MvcResult result = mvc.perform(builder.file(file)).andExpect(MockMvcResultMatchers.status().isOk()).andReturn();
		assertTrue(result.getResponse().getContentAsString().contains("Hello World!"));
	}

	private InputStream getResource(String path) {
		return MicrosoftApplicationTests.class.getResourceAsStream(path);
	}

	private byte[] getResourceAsByte(String path) throws IOException {
		return IOUtils.toByteArray(getResource(path));
	}

}