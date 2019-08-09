package com.excel.json.core.servlets;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.json.JSONArray;

import javax.jcr.Node;
import javax.jcr.NodeIterator;
import javax.jcr.PathNotFoundException;
import javax.jcr.RepositoryException;
import javax.jcr.Session;
import javax.servlet.Servlet;
import javax.servlet.ServletException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.servlets.HttpConstants;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.io.FilenameUtils;
import org.apache.sling.api.request.RequestParameter;
import org.apache.sling.api.resource.LoginException;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.resource.ResourceResolverFactory;
import org.osgi.framework.Constants;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

@Component(service = Servlet.class, property = { Constants.SERVICE_DESCRIPTION + "=Excel to JSON Converter Servlet",
		"sling.servlet.methods=" + HttpConstants.METHOD_POST, "sling.servlet.paths=" + "/apps/get/json/from/xls" })
public class ExcelToJSONServlet extends SlingAllMethodsServlet {

	protected static final String XLSX = "xlsx";
	protected static final String XLS = "xls";
	protected static final String DEFAULT_FILE_EXTENSION = "json";
	protected static final String DEFAULT_FILE_NAME = "renderedJSON." + DEFAULT_FILE_EXTENSION;

	private static final long serialVersionUID = 1L;

	private static final Logger LOGGER = LoggerFactory.getLogger(ExcelToJSONServlet.class);

	// Inject a Sling ResourceResolverFactory
	@Reference
	private ResourceResolverFactory resolverFactory;

	@Override
	protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response)
			throws ServletException, IOException {
		// TODO Auto-generated method stub
		try {
			final boolean isMultipart = ServletFileUpload.isMultipartContent(request);

			if (isMultipart) {

				final Map<String, RequestParameter[]> params = request.getRequestParameterMap();

				Iterator<Map.Entry<String, RequestParameter[]>> parmsIterator = params.entrySet().iterator();

				Map.Entry<String, RequestParameter[]> pairs = parmsIterator.next();
				RequestParameter[] parameterArray = pairs.getValue();
				RequestParameter param = parameterArray[0];

				final InputStream stream = param.getInputStream();

				String mimeType = FilenameUtils.getExtension(param.getFileName());
				Workbook workbook = null;

				if (mimeType.equals(XLSX)) {
					workbook = new XSSFWorkbook(stream);
				}

				if (mimeType.equals(XLS)) {
					workbook = new HSSFWorkbook(stream);
				}

				// The InputStream represents the excel file
				ArrayList<HashMap<String, String>> excelData = readExcel(workbook);

				stream.close();
				workbook.close();

				InputStream jsonStream = null;
				if (excelData != null) {
					JSONArray parsedJSON = new JSONArray(excelData);
					jsonStream = new ByteArrayInputStream(parsedJSON.toString().getBytes());
				}

				pairs = parmsIterator.next();
				parameterArray = pairs.getValue();
				param = parameterArray[0];

				String destPath = param.getString();

				storeFileinJCR(destPath, DEFAULT_FILE_NAME, DEFAULT_FILE_EXTENSION, jsonStream);
				jsonStream.close();

				response.setStatus(200);
				response.getWriter().print("Successfully parsed the file!");

			}
		}

		catch (Exception e) {
			e.printStackTrace();
			response.setStatus(403);
			response.getWriter().print("Please check the inputs!");
		}
	}

	// Read Excel File
	private ArrayList<HashMap<String, String>> readExcel(Workbook workbook) throws IOException {

		ArrayList<HashMap<String, String>> resultList = new ArrayList<HashMap<String, String>>();
		String[] tableHeader = getTableHeader(workbook);

		if (tableHeader != null) {
			Sheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				if (row.getRowNum() == 0) {
					continue;
				}

				HashMap<String, String> rowData = new HashMap<String, String>();
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					rowData.put(tableHeader[cell.getColumnIndex()], cell.toString());
				}

				resultList.add(rowData);
			}
		}

		return resultList;

	}

	// Get table's header / first row
	private String[] getTableHeader(Workbook workbook) throws IOException {

		Sheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();

		Row row = rowIterator.next();
		Iterator<Cell> cellIterator = row.cellIterator();
		String[] tableHeader = new String[row.getLastCellNum()];

		int index = 0;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			tableHeader[index] = cell.toString();
			index++;
		}

		return tableHeader;

	}

	// Store file in JCR
	private void storeFileinJCR(String destPath, String fileName, String mimetype, InputStream jsonStream)
			throws LoginException, PathNotFoundException, RepositoryException {
		Map<String, Object> serviceNameParam = new HashMap<String, Object>();
		serviceNameParam.put(ResourceResolverFactory.SUBSERVICE, "excel-to-json");
		ResourceResolver resolver = resolverFactory.getServiceResourceResolver(serviceNameParam);
		Session session = resolver.adaptTo(Session.class);

		Node node = session.getNode(destPath);
		NodeIterator childNodes = node.getNodes(fileName);

		if (childNodes.hasNext()) {
			childNodes.nextNode().remove();
		}

		javax.jcr.ValueFactory valueFactory = session.getValueFactory();
		javax.jcr.Binary contentValue = valueFactory.createBinary(jsonStream);
		Node fileNode = node.addNode(fileName, "nt:file");
		fileNode.addMixin("mix:referenceable");
		Node resNode = fileNode.addNode("jcr:content", "nt:resource");
		resNode.setProperty("jcr:mimeType", mimetype);
		resNode.setProperty("jcr:data", contentValue);
		Calendar lastModified = Calendar.getInstance();
		lastModified.setTimeInMillis(lastModified.getTimeInMillis());
		resNode.setProperty("jcr:lastModified", lastModified);
		session.save();
		session.logout();

	}

}