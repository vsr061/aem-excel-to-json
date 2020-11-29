package com.excel.json.core.servlets;

import com.day.cq.dam.api.Asset;
import com.day.cq.dam.api.AssetManager;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.request.RequestParameter;
import org.apache.sling.api.resource.LoginException;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.resource.ResourceResolverFactory;
import org.apache.sling.api.servlets.HttpConstants;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.json.JSONArray;
import org.osgi.framework.Constants;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.Servlet;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;

/**
 * @author https://github.com/vsr061
 * This servlet parses an Excel file table stored in a first sheet to JSON array
 * where each object in the array represents a row in the Excel file and object
 * member keys are the table header values.
 * ___________________________________________
 * | first name | last name | profile avatar |
 * -------------------------------------------
 * | Viraj      | Rane      | pacman         |
 * -------------------------------------------
 * | Json       | Jose      | ninja          |
 * -------------------------------------------
 *
 * will get parsed to
 *
 * {
 *     [
 *         {
 *             "first name": "Viraj",
 *             "last name": "Rane",
 *             "profile avatar": "pacman"
 *         },
 *         {
 *            "first name": "Json",
 *            "last name": "Jose",
 *            "profile avatar": "ninja"
 *         }
 *     ]
 * }
 */
@Component(
        service = Servlet.class,
        property =
                {
                        Constants.SERVICE_DESCRIPTION + "=Excel to JSON Converter Servlet",
                        "sling.servlet.methods=" + HttpConstants.METHOD_POST,
                        "sling.servlet.paths=" + "/apps/get/json/from/xls"
                }
)
public class ExcelToJSONServlet extends SlingAllMethodsServlet {

    private static final String XLSX = "xlsx";
    private static final String XLS = "xls";
    private static final String JSON_MIME_TYPE = "application/json";
    private static final String JSON_FILE_EXTENSION = ".json";
    private static final String SUB_SERVICE_NAME = "excel-to-json";

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelToJSONServlet.class);

    @Reference
    protected ResourceResolverFactory resolverFactory;

    @Override
    protected void doPost(SlingHttpServletRequest request, SlingHttpServletResponse response)
            throws IOException {

        final boolean isMultipart = ServletFileUpload.isMultipartContent(request);
        final Map<String, RequestParameter[]> params = request.getRequestParameterMap();
        Iterator<Map.Entry<String, RequestParameter[]>> paramsIterator = params.entrySet().iterator();

        Map.Entry<String, RequestParameter[]> pairs = paramsIterator.next();
        RequestParameter[] parameterArray = pairs.getValue();

        //Uploaded file
        RequestParameter file = parameterArray[0];

        //File's mime type
        String mimeType = FilenameUtils.getExtension(file.getFileName());

        pairs = paramsIterator.next();
        parameterArray = pairs.getValue();
        RequestParameter path = parameterArray[0];

        //JCR path to store the parsed JSON as an asset
        String destinationPath = path.getString();

        if (isMultipart
                && StringUtils.isNotEmpty(mimeType)
                && mimeType.contains(XLS)
                && StringUtils.isNotBlank(destinationPath)
        ) {
            Workbook workbook = null;
            InputStream jsonStream = null;
            try (final InputStream stream = file.getInputStream()) {

                if (Objects.nonNull(stream)) {

                    if (mimeType.equals(XLSX)) {
                        workbook = new XSSFWorkbook(stream);
                    } else if (mimeType.equals(XLS)) {
                        workbook = new HSSFWorkbook(stream);
                    } else {
                        LOGGER.error("{}: Unsupported file type!", mimeType);
                        sendStatus(
                                response,
                                SlingHttpServletResponse.SC_BAD_REQUEST,
                                mimeType.concat(" Unsupported file type!")
                        );
                        return;
                    }

                    //Get list of map objects
                    ArrayList<HashMap<String, String>> excelData = readExcel(workbook);

                    //Parse list of map objects to JSON array
                    JSONArray parsedJSON = new JSONArray(excelData);

                    /*
                      Convert JSON array to stream
                      ** Keep encoding as UTF-8 **
                    */
                    jsonStream = new ByteArrayInputStream(parsedJSON.toString().getBytes(StandardCharsets.UTF_8));

                    //store stream in JCR as AEM asset
                    Asset fileInJCR =
                            storeFileInJCR(
                                    destinationPath,
                                    FilenameUtils.getBaseName(file.getFileName()).concat(JSON_FILE_EXTENSION),
                                    jsonStream);

                    int statusCode = Objects.nonNull(fileInJCR)
                            ? SlingHttpServletResponse.SC_OK
                            : SlingHttpServletResponse.SC_INTERNAL_SERVER_ERROR;
                    String message = Objects.nonNull(fileInJCR)
                            ? "Successfully parsed the file!"
                            : "Error occurred while saving the file! Check system user permissions";
                    sendStatus(response, statusCode, message);

                } else {
                    sendStatus(
                            response,
                            SlingHttpServletResponse.SC_BAD_REQUEST,
                            "Empty file!"
                    );
                }
            } catch (LoginException e) {
                LOGGER.error("Error Occurred: {}", e.getMessage());
                sendStatus(
                        response,
                        SlingHttpServletResponse.SC_INTERNAL_SERVER_ERROR,
                        "Error occurred while saving the file!"
                );
            } finally {
                if (Objects.nonNull(workbook) && Objects.nonNull(jsonStream)) {
                    workbook.close();
                    jsonStream.close();
                }
            }
        } else {
            sendStatus(
                    response,
                    SlingHttpServletResponse.SC_BAD_REQUEST,
                    "Unsupported request payload!"
            );
        }
    }

    private void sendStatus(SlingHttpServletResponse response, int status, String message) throws IOException {
        response.setCharacterEncoding(StandardCharsets.UTF_8.displayName());
        response.setContentType("text/plain");
        response.setStatus(status);
        response.getWriter().print(message);
    }

    // Read Excel File
    private ArrayList<HashMap<String, String>> readExcel(Workbook workbook) {
        ArrayList<HashMap<String, String>> resultList = new ArrayList<>();
        String[] tableHeader = getTableHeader(workbook);
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }
            HashMap<String, String> rowData = new HashMap<>();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                rowData.put(tableHeader[cell.getColumnIndex()], cell.getRichStringCellValue().getString());
            }
            if (!rowData.isEmpty()) {
                resultList.add(rowData);
            }
        }
        return resultList;
    }

    // Get table's header / first row
    private String[] getTableHeader(Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next();
        Iterator<Cell> cellIterator = row.cellIterator();
        String[] tableHeader = new String[row.getLastCellNum()];
        for (int index = 0; cellIterator.hasNext(); index++) {
            Cell cell = cellIterator.next();
            tableHeader[index] = cell.getRichStringCellValue().getString();
        }
        return tableHeader;
    }

    // Store file in JCR
    private Asset storeFileInJCR(String destinationPath, String fileName, InputStream jsonStream)
            throws LoginException {
        Map<String, Object> serviceNameParam = new HashMap<>();
        serviceNameParam.put(ResourceResolverFactory.SUBSERVICE, SUB_SERVICE_NAME);
        ResourceResolver resolver = resolverFactory.getServiceResourceResolver(serviceNameParam);
        AssetManager assetManager = resolver.adaptTo(AssetManager.class);
        if (Objects.nonNull(assetManager)) {
            return assetManager.createAsset(
                    destinationPath.concat("/".concat(fileName)),
                    jsonStream,
                    JSON_MIME_TYPE,
                    true);
        } else {
            return null;
        }
    }
}
