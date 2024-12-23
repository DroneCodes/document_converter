package org.fisayo;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * The DocumentConverter class provides methods to convert documents between different formats.
 * Supported conversions include:
 * - Excel to JSON
 * - Access to JSON
 * - JSON to Excel
 * - JSON to Access
 */
public class DocumentConverter {
    private static final String ASSETS_DIR = "assets";
    private static final String RESULTS_DIR = "results";
    private static final ObjectMapper mapper = new ObjectMapper();

    /**
     * The main method that runs the document converter application.
     * It provides a menu for the user to select the type of conversion and handles the conversion process.
     *
     * @param args Command line arguments (not used).
     */
    public static void main(String[] args) {
        // Check if assets directory exists
        if (!Files.exists(Paths.get(ASSETS_DIR))) {
            System.err.println("Error: 'assets' directory not found!");
            System.out.println("Please create an 'assets' directory and add your files there.");
            return;
        }

        // Create results directory if it doesn't exist
        try {
            Files.createDirectories(Paths.get(RESULTS_DIR));
        } catch (IOException e) {
            System.err.println("Error creating results directory: " + e.getMessage());
            return;
        }

        Scanner scanner = new Scanner(System.in);
        while (true) {
            System.out.println("\nDocument Converter Menu:");
            System.out.println("1. Convert Excel to JSON");
            System.out.println("2. Convert Access to JSON");
            System.out.println("3. Convert JSON to Excel");
            System.out.println("4. Convert JSON to Access");
            System.out.println("5. Exit");

            System.out.print("Enter your choice (1-5): ");
            int choice = scanner.nextInt();
            scanner.nextLine(); // Consume newline

            if (choice == 5) break;

            // Get list of files based on selected conversion type
            List<String> availableFiles = listAvailableFiles(choice);
            if (availableFiles.isEmpty()) {
                System.out.println("No compatible files found in the assets directory!");
                continue;
            }

            // Display available files
            System.out.println("\nAvailable files in assets directory:");
            for (int i = 0; i < availableFiles.size(); i++) {
                System.out.println((i + 1) + ". " + availableFiles.get(i));
            }

            // Get user's file selection
            System.out.print("Select file number: ");
            int fileChoice = scanner.nextInt();
            scanner.nextLine(); // Consume newline

            if (fileChoice < 1 || fileChoice > availableFiles.size()) {
                System.out.println("Invalid file selection!");
                continue;
            }

            String inputFile = availableFiles.get(fileChoice - 1);

            // Get output filename
            System.out.print("Enter output filename (with appropriate extension): ");
            String outputFile = scanner.nextLine();

            try {
                switch (choice) {
                    case 1:
                        excelToJson(inputFile, outputFile);
                        break;
                    case 2:
                        accessToJson(inputFile, outputFile);
                        break;
                    case 3:
                        jsonToExcel(inputFile, outputFile);
                        break;
                    case 4:
                        jsonToAccess(inputFile, outputFile);
                        break;
                }
                System.out.println("Conversion completed successfully!");
                System.out.println("Output file saved as: " + RESULTS_DIR + "/" + outputFile);
            } catch (Exception e) {
                System.err.println("Error during conversion: " + e.getMessage());
            }
        }
    }

    /**
     * Lists available files in the assets directory based on the selected conversion type.
     *
     * @param conversionType The type of conversion selected by the user.
     * @return A list of available files for the selected conversion type.
     */
    private static List<String> listAvailableFiles(int conversionType) {
        List<String> files = new ArrayList<>();
        File assetsDir = new File(ASSETS_DIR);
        File[] allFiles = assetsDir.listFiles();

        if (allFiles == null) return files;

        for (File file : allFiles) {
            String fileName = file.getName().toLowerCase();
            switch (conversionType) {
                case 1: // Excel to JSON
                    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
                        files.add(file.getName());
                    }
                    break;
                case 2: // Access to JSON
                    if (fileName.endsWith(".accdb") || fileName.endsWith(".mdb")) {
                        files.add(file.getName());
                    }
                    break;
                case 3: // JSON to Excel
                case 4: // JSON to Access
                    if (fileName.endsWith(".json")) {
                        files.add(file.getName());
                    }
                    break;
            }
        }
        return files;
    }

    /**
     * Converts an Excel file to a JSON file.
     *
     * @param inputFile  The name of the input Excel file.
     * @param outputFile The name of the output JSON file.
     * @throws IOException If an I/O error occurs.
     */
    private static void excelToJson(String inputFile, String outputFile) throws IOException {
        Workbook workbook = WorkbookFactory.create(new File(ASSETS_DIR + "/" + inputFile));
        Sheet sheet = workbook.getSheetAt(0);

        Row headerRow = sheet.getRow(0);
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }

        ArrayNode jsonArray = mapper.createArrayNode();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            ObjectNode jsonRow = mapper.createObjectNode();

            for (int j = 0; j < headers.size(); j++) {
                Cell cell = row.getCell(j);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            jsonRow.put(headers.get(j), cell.getStringCellValue());
                            break;
                        case NUMERIC:
                            jsonRow.put(headers.get(j), cell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            jsonRow.put(headers.get(j), cell.getBooleanCellValue());
                            break;
                    }
                }
            }
            jsonArray.add(jsonRow);
        }

        mapper.writerWithDefaultPrettyPrinter()
                .writeValue(new File(RESULTS_DIR + "/" + outputFile), jsonArray);

        workbook.close();
    }

    /**
     * Converts an Access database file to a JSON file.
     *
     * @param inputFile  The name of the input Access database file.
     * @param outputFile The name of the output JSON file.
     * @throws Exception If an error occurs during the conversion.
     */
    private static void accessToJson(String inputFile, String outputFile) throws Exception {
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
        String dbURL = "jdbc:ucanaccess://" + ASSETS_DIR + "/" + inputFile;

        try (Connection conn = DriverManager.getConnection(dbURL)) {
            DatabaseMetaData meta = conn.getMetaData();
            ResultSet tables = meta.getTables(null, null, "%", new String[]{"TABLE"});

            ObjectNode dbJson = mapper.createObjectNode();

            while (tables.next()) {
                String tableName = tables.getString("TABLE_NAME");
                Statement stmt = conn.createStatement();
                ResultSet rs = stmt.executeQuery("SELECT * FROM " + tableName);

                ArrayNode tableData = mapper.createArrayNode();
                ResultSetMetaData rsmd = rs.getMetaData();
                int columnCount = rsmd.getColumnCount();

                while (rs.next()) {
                    ObjectNode row = mapper.createObjectNode();
                    for (int i = 1; i <= columnCount; i++) {
                        String columnName = rsmd.getColumnName(i);
                        Object value = rs.getObject(i);
                        if (value != null) {
                            if (value instanceof Number) {
                                row.put(columnName, ((Number) value).doubleValue());
                            } else {
                                row.put(columnName, value.toString());
                            }
                        }
                    }
                    tableData.add(row);
                }
                dbJson.set(tableName, tableData);
            }

            mapper.writerWithDefaultPrettyPrinter()
                    .writeValue(new File(RESULTS_DIR + "/" + outputFile), dbJson);
        }
    }

    /**
     * Converts a JSON file to an Excel file.
     *
     * @param inputFile  The name of the input JSON file.
     * @param outputFile The name of the output Excel file.
     * @throws IOException If an I/O error occurs.
     */
    private static void jsonToExcel(String inputFile, String outputFile) throws IOException {
        ArrayNode jsonArray = (ArrayNode) mapper.readTree(new File(ASSETS_DIR + "/" + inputFile));

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Data");

        // Create header row
        if (jsonArray.size() > 0) {
            Row headerRow = sheet.createRow(0);
            ObjectNode firstRow = (ObjectNode) jsonArray.get(0);
            Iterator<String> fieldNames = firstRow.fieldNames();
            int colIdx = 0;
            while (fieldNames.hasNext()) {
                String fieldName = fieldNames.next();
                headerRow.createCell(colIdx).setCellValue(fieldName);
                colIdx++;
            }
        }

        // Create data rows
        for (int i = 0; i < jsonArray.size(); i++) {
            ObjectNode jsonRow = (ObjectNode) jsonArray.get(i);
            Row row = sheet.createRow(i + 1);

            Iterator<Map.Entry<String, JsonNode>> fields = jsonRow.fields();
            int colIdx = 0;
            while (fields.hasNext()) {
                Map.Entry<String, JsonNode> field = fields.next();
                Cell cell = row.createCell(colIdx);

                JsonNode value = field.getValue();
                if (value.isNumber()) {
                    cell.setCellValue(value.asDouble());
                } else if (value.isBoolean()) {
                    cell.setCellValue(value.asBoolean());
                } else {
                    cell.setCellValue(value.asText());
                }
                colIdx++;
            }
        }

        // Auto-size columns
        for (int i = 0; i < sheet.getRow(0).getLastCellNum(); i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream fileOut = new FileOutputStream(RESULTS_DIR + "/" + outputFile)) {
            workbook.write(fileOut);
        }
        workbook.close();
    }

    /**
     * Converts a JSON file to an Access database file.
     *
     * @param inputFile  The name of the input JSON file.
     * @param outputFile The name of the output Access database file.
     * @throws Exception If an error occurs during the conversion.
     */

    private static void jsonToAccess(String inputFile, String outputFile) throws Exception {
        // Read the JSON file
        JsonNode rootNode = mapper.readTree(new File(ASSETS_DIR + "/" + inputFile));
        Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");

        // Create new Access database
        String dbURL = "jdbc:ucanaccess://" + RESULTS_DIR + "/" + outputFile;
        try (Connection conn = DriverManager.getConnection(dbURL + ";newDatabaseVersion=V2010")) {
            if (rootNode.isArray()) {
                // Handle array type JSON
                ArrayNode arrayNode = (ArrayNode) rootNode;
                if (arrayNode.size() > 0) {
                    // Create single table for array data
                    String tableName = "JsonData";
                    ObjectNode firstRow = (ObjectNode) arrayNode.get(0);

                    // Create table
                    StringBuilder createTableSQL = new StringBuilder();
                    createTableSQL.append("CREATE TABLE ").append(tableName).append(" (");

                    Iterator<String> columns = firstRow.fieldNames();
                    List<String> columnNames = new ArrayList<>();
                    while (columns.hasNext()) {
                        String column = columns.next();
                        columnNames.add(column);
                        JsonNode value = firstRow.get(column);
                        String type = determineAccessType(value);
                        createTableSQL.append(column).append(" ").append(type);
                        if (columns.hasNext()) createTableSQL.append(", ");
                    }
                    createTableSQL.append(")");

                    System.out.println("Creating table with SQL: " + createTableSQL);
                    Statement stmt = conn.createStatement();
                    stmt.executeUpdate(createTableSQL.toString());

                    // Insert data
                    for (JsonNode row : arrayNode) {
                        StringBuilder insertSQL = new StringBuilder();
                        insertSQL.append("INSERT INTO ").append(tableName).append(" (");
                        insertSQL.append(String.join(", ", columnNames));
                        insertSQL.append(") VALUES (");

                        // Create placeholders for PreparedStatement
                        String placeholders = "?" + ", ?".repeat(columnNames.size() - 1);
                        insertSQL.append(placeholders).append(")");

                        PreparedStatement pstmt = conn.prepareStatement(insertSQL.toString());

                        // Set values in PreparedStatement
                        for (int i = 0; i < columnNames.size(); i++) {
                            JsonNode value = row.get(columnNames.get(i));
                            if (value.isNull()) {
                                pstmt.setNull(i + 1, Types.VARCHAR);
                            } else if (value.isNumber()) {
                                pstmt.setDouble(i + 1, value.asDouble());
                            } else if (value.isBoolean()) {
                                pstmt.setBoolean(i + 1, value.asBoolean());
                            } else {
                                pstmt.setString(i + 1, value.asText());
                            }
                        }

                        pstmt.executeUpdate();
                    }
                }
            } else if (rootNode.isObject()) {
                // Original code for handling object type JSON
                ObjectNode dbJson = (ObjectNode) rootNode;
                Iterator<Map.Entry<String, JsonNode>> tables = dbJson.fields();

                while (tables.hasNext()) {
                    Map.Entry<String, JsonNode> table = tables.next();
                    String tableName = table.getKey();
                    ArrayNode tableData = (ArrayNode) table.getValue();

                    if (tableData.size() > 0) {
                        // Create table
                        ObjectNode firstRow = (ObjectNode) tableData.get(0);
                        StringBuilder createTableSQL = new StringBuilder();
                        createTableSQL.append("CREATE TABLE ").append(tableName).append(" (");

                        Iterator<String> columns = firstRow.fieldNames();
                        while (columns.hasNext()) {
                            String column = columns.next();
                            JsonNode value = firstRow.get(column);
                            String type = determineAccessType(value);
                            createTableSQL.append(column).append(" ").append(type);
                            if (columns.hasNext()) createTableSQL.append(", ");
                        }
                        createTableSQL.append(")");

                        Statement stmt = conn.createStatement();
                        stmt.executeUpdate(createTableSQL.toString());

                        // Insert data
                        for (JsonNode row : tableData) {
                            StringBuilder insertSQL = new StringBuilder();
                            insertSQL.append("INSERT INTO ").append(tableName).append(" (");

                            ObjectNode objRow = (ObjectNode) row;
                            Iterator<String> fieldNames = objRow.fieldNames();
                            StringBuilder values = new StringBuilder();
                            List<Object> params = new ArrayList<>();

                            while (fieldNames.hasNext()) {
                                String field = fieldNames.next();
                                insertSQL.append(field);
                                values.append("?");
                                if (fieldNames.hasNext()) {
                                    insertSQL.append(", ");
                                    values.append(", ");
                                }
                                params.add(objRow.get(field).asText());
                            }

                            insertSQL.append(") VALUES (").append(values).append(")");

                            PreparedStatement pstmt = conn.prepareStatement(insertSQL.toString());
                            for (int i = 0; i < params.size(); i++) {
                                pstmt.setObject(i + 1, params.get(i));
                            }
                            pstmt.executeUpdate();
                        }
                    }
                }
            }
        }
        System.out.println("Successfully created Access database: " + RESULTS_DIR + "/" + outputFile);
    }

    /**
     * Determines the appropriate Access database column type for a given JSON value.
     *
     * @param value The JSON value.
     * @return The corresponding Access database column type.
     */
    private static String determineAccessType(JsonNode value) {
        if (value.isNumber()) {
            return "DOUBLE";
        } else if (value.isBoolean()) {
            return "BOOLEAN";
        } else {
            return "TEXT";
        }
    }
}