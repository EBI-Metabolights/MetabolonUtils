package uk.ac.ebi.metabolights.utils.metabolonutils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.isatools.isacreator.configuration.io.ConfigXMLParser;
import org.isatools.isacreator.spreadsheet.model.TableReferenceObject;
import org.isatools.isatab.configurator.schema.IsaTabConfigurationType;
import org.isatools.isatab.configurator.schema.IsatabConfigFileDocument;
import org.isatools.plugins.metabolights.assignments.actions.AutoCompletionAction;
import org.isatools.plugins.metabolights.assignments.model.Metabolite;

import java.io.*;
import java.util.ArrayList;
import java.util.Vector;
import java.util.concurrent.atomic.AtomicReference;

public class FileUtils {

    private static String configPath = "." + File.separator + "metabolomics_configuration" + File.separator;
    private static String configFile = configPath + "configuration_ms.xml";
    private static String configurationFile = FileUtils.class.getClassLoader().getResource(configFile).getFile();

    private final static String sampleNameAnnotation = "SAMPLE_NAME";
    private final static String clientIdAnnotation = "CLIENT_IDENTIFIER";
    private final static String parentSampleIdAnnotation = "PARENT_SAMPLE_ID";
    private final static String startingVolumeAnnotation = "STARTING_VOLUME";
    private final static String headersAnnotation = "HEADERS";
    private final static String dataAnnotation = "DATA";
    private final static String MAFSheetName = "MAF Data";
    private final static String metabolonSheetName = "Metabolon Data";
    private final static String annotatedSheetName = "Annotated Data";
    private static int lastCellNumber = 0;

    public static int getLastCellNumber() {
        return lastCellNumber;
    }

    public static void setLastCellNumber(int lastCellNumber) {
        FileUtils.lastCellNumber = lastCellNumber;
    }

    public TableReferenceObject getMSConfig() {
        return getConfiguration(configurationFile);
    }

    public void convertExcelFile(String fileName) throws IOException, InvalidFormatException {
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(fileName));
        workbook.setSheetName(workbook.getSheetIndex(workbook.getSheetAt(0)), metabolonSheetName);
        workbook.cloneSheet(0); //Copy the existing sheet

        //Retrieving the cloned worksheet
        Sheet metabolonSheet = workbook.getSheetAt(1);   //Sheet 0 is the original data from Metabolon
        workbook.setSheetName(workbook.getSheetIndex(metabolonSheet), annotatedSheetName);

        //Annotation rows in the original Metabolon sheet
        metabolonSheet.forEach(row -> {
            annotateMetabolonData(row, row.getRowNum());
        });

        // Create the new sheet for MAF
        Sheet newSheet = addStandardHeaderRow(workbook);
        newSheet = addSampleHeaderRow(newSheet, metabolonSheet);
        addMetabolonData(newSheet, metabolonSheet);

        // Write the output to a new Excel file
        FileOutputStream fileOut = new FileOutputStream("MetabolonPeakAreaTable_MAF.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    private Sheet addStandardHeaderRow(Workbook workbook){
        Sheet newSheet = workbook.createSheet(MAFSheetName);
        Row headerRow = newSheet.createRow(0);
        TableReferenceObject mafTable = getMSConfig();  //Header values from the config file

        Vector<String> standardHeaders =  mafTable.getHeaders(); //Get all the headers from the config file
        for (int i = 0; i < standardHeaders.size(); i++) {
            if (i>0)  //Skip the first row as this only has the row-number    ("Row No.")
                headerRow.createCell(i-1).setCellValue(standardHeaders.get(i));
        }

        return newSheet;

    }

    private Sheet addSampleHeaderRow(Sheet newSheet, Sheet metabolonSheet){
        Row headerRow = newSheet.getRow(0); //The header row is at the top of the sheet
        ArrayList<String> sampleColumns = new ArrayList<String>();

        short lastCell = headerRow.getLastCellNum();
        lastCell--;   //Starts at 1!
        setLastCellNumber(lastCell);

        metabolonSheet.forEach(row -> {
            Row currentRow = row;
            if (currentRow.getCell(0).getRichStringCellValue().toString().equals(sampleNameAnnotation)) {
                currentRow.forEach(cell -> {
                    String cellValue = "";

                    if (cell.getRichStringCellValue().getString() != null)
                        cellValue = cell.getRichStringCellValue().getString();

                    if (cellValue.length() > 1 && !cellValue.equals(sampleNameAnnotation))
                        sampleColumns.add(cellValue);
                });
            }
        });

        //Add sample rows at the end of the header row
        for (int i = 0; i < sampleColumns.size(); i++) {
            String sampleName = sampleColumns.get(i);
            headerRow.createCell(lastCell+i).setCellValue(sampleName);
        }

        return newSheet;

    }

    private Metabolite getMetaboliteInformation(String identifier, String metabolite){
        // search by compound name
        Metabolite met = new Metabolite();

        if (identifier != null)
            met = AutoCompletionAction.getMetaboliteFromMetaboLightWS(AutoCompletionAction.IDENTIFIER_COL_NAME, identifier);

        if (metabolite != null)
            met = AutoCompletionAction.getMetaboliteFromMetaboLightWS(AutoCompletionAction.DESCRIPTION_COL_NAME, metabolite);

        return met;
    }

    private void addMetabolonData(Sheet newSheet, Sheet metabolonSheet){

        //Output the header row first
        newSheet.getRow(0).forEach(newCell -> {
            printCellValue(newCell);
        });
        System.out.println(); //To ensure line breaks when printing to screen

        metabolonSheet.forEach(row -> {
            String dataRowType = row.getCell(0).getRichStringCellValue().toString();

            //Retrieving cells
            if (dataRowType.equals(dataAnnotation)){

                int numberOfCells = (int) newSheet.getRow(0).getLastCellNum();  // Get the header row
                int newRowNum = newSheet.getLastRowNum();;
                Row newRow = newSheet.createRow(newRowNum+1); //Don't overwrite the last or header row.
                for (int i = 0; i < numberOfCells; i++) {
                    newRow.createCell(i).setCellValue("");   //Add all the empty cells first
                }


                //Lamda requirement?
                final String[] dbId = { null };
                final String[] metName = { null };

                row.forEach(cell -> {
                    int metabolonColumnNumber = cell.getColumnIndex();

                    if (metabolonColumnNumber == 11 || metabolonColumnNumber == 12) { // KEGG (11) or HMDB (12)
                        if (cell.getRichStringCellValue().getString() != null && cell.getRichStringCellValue().getString().length() > 2)
                            dbId[0] = cell.getRichStringCellValue().getString(); // to be used for "database_identifier".
                        // Adding KEGG first, in case HMDB is not reported
                    }

                    if (dbId[0] != null && dbId[0].length() > 2)
                        newRow.createCell(0).setCellValue(dbId[0]); // "database_identifier"

                    if (metabolonColumnNumber == 1) { //Compound name
                        metName[0] = cell.getRichStringCellValue().getString();
                        newRow.createCell(4).setCellValue(metName[0]); // "metabolite_identification"
                    }

                    if (metabolonColumnNumber == 8) { //Mass
                        Double cellValue = cell.getNumericCellValue();
                        if (cellValue != null)
                            newRow.createCell(5).setCellValue(cell.getNumericCellValue()); // "mass_to_charge"
                    }

                    if (metabolonColumnNumber >= 13) {  // The sample  concentration values starts at column 13
                        Double cellValue = cell.getNumericCellValue();
                        if (cellValue != null)
                            newRow.createCell((getLastCellNumber()+ metabolonColumnNumber)-13 ).setCellValue(cellValue); // Sample data
                        //TODO, fix the -13 hack!
                    }


                });

                //if (dbId[0] != null && dbId[0].length() > 2)
                if (metName[0] != null) {

                    String cleanMetName = metName[0].replaceAll("\\*","");
                    Metabolite met = getMetaboliteInformation(dbId[0], cleanMetName);

                    if (met != null) { // Replace with MetaboLights WS search results
                        String identifier, formula, description, smiles, inchi;
                        identifier = met.getIdentifier();
                        formula = met.getFormula();
                        description = met.getDescription();
                        smiles = met.getSmiles();
                        inchi = met.getInchi();

                        if (identifier != null)
                            newRow.createCell(0).setCellValue(identifier);

                        if (formula != null)
                            newRow.createCell(1).setCellValue(formula);

                        if (smiles != null)
                            newRow.createCell(2).setCellValue(smiles);

                        if (inchi != null)
                            newRow.createCell(3).setCellValue(inchi);

                        if (description != null)
                            newRow.createCell(4).setCellValue(description);

                    }
                }


                //Printing new value to cell
                if (newRow != null)
                    newRow.forEach(newCell -> {
                        printCellValue(newCell);
                    });

                System.out.println(); //To ensure line breaks when printing to screen

            }
        });

    }

    /**
     * Adds a cell in the first row of the spreadsheet to indicate what type of data the row consists of
     * @param row
     * @param rowNum
     */
    private void annotateMetabolonData(Row row, int rowNum){

        AtomicReference<String> cellValue = new AtomicReference<>(dataAnnotation);

        //Adding new value to cell
        row.forEach(cell -> {
            switch (rowNum) {
                case 0:  cellValue.set(clientIdAnnotation); break;
                case 1:  cellValue.set(parentSampleIdAnnotation); break;
                case 2:  cellValue.set(sampleNameAnnotation); break;
                case 3:  cellValue.set(startingVolumeAnnotation); break;
                case 4:  cellValue.set(headersAnnotation); break;
                default: cellValue.set(dataAnnotation);
            }
            row.createCell(0).setCellValue(cellValue.get());  //Change the first cell to our annotation
        });
    }

    private static void printCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
            case STRING: System.out.print(cell.getRichStringCellValue().getString()); break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA: System.out.print(cell.getCellFormula()); break;
            case BLANK: System.out.print(""); break;
            default: System.out.print("");
        }

        System.out.print("\t");
    }


    private TableReferenceObject getConfiguration(String fileName){
        TableReferenceObject tableReferenceObject = null;

        //Load the current settings file
        try {
            InputStream inputStream = new FileInputStream(configurationFile);
            IsatabConfigFileDocument configurationFile = IsatabConfigFileDocument.Factory.parse(inputStream);

            ConfigXMLParser parser = new ConfigXMLParser("");

            //Add columns defined in the configuration file
            for (IsaTabConfigurationType doc : configurationFile.getIsatabConfigFile().getIsatabConfigurationArray()) {
                parser.processTable(doc);
            }

            if (parser.getTables().size() > 0) {
                return parser.getTables().get(0);
            }
        } catch (XmlException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return tableReferenceObject;
    }


}
