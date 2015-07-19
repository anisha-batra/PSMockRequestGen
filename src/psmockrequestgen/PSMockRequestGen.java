/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package psmockrequestgen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author anisha
 */
public class PSMockRequestGen {

    static int COLID_URL = 0;
    static int COLID_OUTPUTFILE = 1;
    static int COLID_RESPONSEFILE = 2;
    static int COLID_SCENARIONAME = 3;
    static int COLID_REQUIREDSCENARIOSTATE = 4;
    static int COLID_NEWSCENARIOSTATE = 5;
    static int COLID_BODYCONTAINS = 6;

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException {
        String excelFileName = "";
        String ouputBaseFolder = "";

        /* Set Default Values */
        excelFileName = "C:\\Users\\anisha\\Desktop\\Anisha\\RequestGen\\Requests.xlsx";
        ouputBaseFolder = "C:/mappings";

        /* Read Command Line Arguments */
        if (args.length > 0) {
            excelFileName = args[0];
        }

        if (args.length > 1) {
            ouputBaseFolder = args[1];
        }

        /* Open Excel File */
        InputStream excelFileToRead = new FileInputStream(excelFileName);
        XSSFWorkbook wb = new XSSFWorkbook(excelFileToRead);
        XSSFSheet sheet = wb.getSheetAt(0);

        int totalNumOfDataRows = sheet.getPhysicalNumberOfRows() - 1;

        for (int rowIndex = 1; rowIndex <= totalNumOfDataRows; rowIndex++) {
            
            try {
                String urlPath = getCellValue(sheet, rowIndex, COLID_URL);
                String outputFile = ouputBaseFolder + "/" + getCellValue(sheet, rowIndex, COLID_OUTPUTFILE);
                String responseFile = getCellValue(sheet, rowIndex, COLID_RESPONSEFILE);
                String scenarioName = getCellValue(sheet, rowIndex, COLID_SCENARIONAME);
                String requiredScenarioState = getCellValue(sheet, rowIndex, COLID_REQUIREDSCENARIOSTATE);
                String newScenarioState = getCellValue(sheet, rowIndex, COLID_NEWSCENARIOSTATE);
                String bodyContains = getCellValue(sheet, rowIndex, COLID_BODYCONTAINS);

                ArrayList<String> listOfBodyContains = new ArrayList<String>();
                for (String bodyContainsLine : bodyContains.split("\n")) {
                    if (!bodyContainsLine.equalsIgnoreCase("")) {
                        listOfBodyContains.add(bodyContainsLine);
                    }
                }

                File file = new File(outputFile);

                // if file doesnt exists, then create it
                if (!file.exists()) {
                    file.getParentFile().mkdirs();
                    file.createNewFile();
                }

                PrintWriter printWriter = new PrintWriter(outputFile);

                printWriter.println(String.format("{"));
                if (!scenarioName.equalsIgnoreCase("")) {
                    printWriter.println(String.format("    \"scenarioName\": \"%s\",", scenarioName));
                }
                if (!requiredScenarioState.equalsIgnoreCase("")) {
                    printWriter.println(String.format("    \"requiredScenarioState\": \"%s\",", requiredScenarioState));
                }
                if (!newScenarioState.equalsIgnoreCase("")) {
                    printWriter.println(String.format("    \"newScenarioState\": \"%s\",", newScenarioState));
                }
                printWriter.println(String.format("    \"request\": {"));
                printWriter.println(String.format("        \"method\": \"POST\","));
                printWriter.println(String.format("        \"urlPath\": \"%s\",", urlPath));
                printWriter.println(String.format("        \"bodyPatterns\": ["));
                //printWriter.println(String.format("            {"));
                for (String bodyContainsString : listOfBodyContains) {
                    // Write , if not last string
                    if (bodyContainsString != listOfBodyContains.get(listOfBodyContains.size() - 1)) {
                        printWriter.println(String.format("              { \"contains\": \"%s\" },", bodyContainsString));
                    } else {
                        printWriter.println(String.format("              { \"contains\": \"%s\" }", bodyContainsString));
                    }
                }
                //printWriter.println(String.format("            }"));
                printWriter.println(String.format("        ]"));
                printWriter.println(String.format("    },"));
                printWriter.println(String.format("    \"response\": {"));
                printWriter.println(String.format("        \"status\": 200,"));
                printWriter.println(String.format("        \"bodyFileName\": \"%s\",", responseFile));
                printWriter.println(String.format("        \"headers\": {"));
                printWriter.println(String.format("            \"Access-Control-Allow-Origin\": \"http://localhost:4503\","));
                printWriter.println(String.format("            \"Access-Control-Allow-Credentials\": \"true\","));
                printWriter.println(String.format("            \"Access-Control-Allow-Headers\": \"x-requested-with, Origin, Content-Type, Accept, access-control-allow-headers, authToken, authorization\","));
                printWriter.println(String.format("            \"Access-Control-Allow-Methods\": \"GET, POST, PUT, DELETE, OPTIONS, HEAD\","));
                printWriter.println(String.format("            \"Access-Control-Max-Age\": \"360\","));
                printWriter.println(String.format("            \"Allow\": \"OPTIONS,POST\""));
                printWriter.println(String.format("        }"));
                printWriter.println(String.format("    }"));
                printWriter.println(String.format("}"));

                printWriter.close();

                //System.out.println(url);
                //System.out.println(outputFile);
                //System.out.println(responseFile);
                //System.out.println(bodyContains);
            } catch (Exception ex) {
                System.out.println(String.format("Error: An Error Occurred On Excel Row: %d", rowIndex));
            }
        }
    }

    private static String getCellValue(XSSFSheet sheet, int rowIndex, int colIndex) {
        String value = "";

        try {
            value = sheet.getRow(rowIndex).getCell(colIndex).getStringCellValue();
        } catch (Exception ex) {
            // Do Nothing
        }

        return value;
    }
}
