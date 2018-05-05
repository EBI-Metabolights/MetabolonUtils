package uk.ac.ebi.metabolights.utils.metabolonutils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.isatools.isacreator.configuration.io.ConfigXMLParser;
import org.isatools.isacreator.spreadsheet.model.TableReferenceObject;
import org.isatools.isatab.configurator.schema.IsaTabConfigurationType;
import org.isatools.isatab.configurator.schema.IsatabConfigFileDocument;

import java.io.*;
import java.util.ArrayList;
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
    private final static String sheetName = "MAF";
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

        //Retrieving worksheets
        Sheet metabolonSheet = workbook.getSheetAt(0);

        //Annotation rows in the original Metabolon sheet
        metabolonSheet.forEach(row -> {
            annotateMetabolonData(row, row.getRowNum());
        });

        // Create the new sheet for MAF
        Sheet newSheet = addStandardHeaderRow(workbook);
        newSheet = addSampleHeaderRow(newSheet, metabolonSheet);
        addMetabolonData(newSheet, metabolonSheet);


        // Write the output to a file

        FileOutputStream fileOut = new FileOutputStream("MetabolonPeakAreaTable_MAF.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    private Sheet addStandardHeaderRow(Workbook workbook){
        Sheet newSheet = workbook.createSheet(sheetName + workbook.getNumberOfSheets());
        TableReferenceObject mafTable = getMSConfig();
        Row headerRow = newSheet.createRow(0);

        //TODO, read config from TableReferenceObject
        String[] standardColums = {
                "database_identifier", "chemical_formula", "smiles", "inchi", "metabolite_identification", "mass_to_charge", "fragmentation",
                "modifications", "charge", "retention_time", "taxid", "species", "database", "database_version", "reliability", "uri", "search_engine",
                "search_engine", "search_engine_score", "smallmolecule_abundance_sub", "smallmolecule_abundance_stdev_sub", "smallmolecule_abundance_std_error_sub" };

        for (int i = 0; i < standardColums.length; i++) {
            headerRow.createCell(i).setCellValue(standardColums[i]);
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

                row.forEach(cell -> {
                    int metabolonColumnNumber = cell.getColumnIndex();

                    String dbId = "";

                    if (metabolonColumnNumber == 11) { // KEGG
                        if (cell.getRichStringCellValue().getString() != null && cell.getRichStringCellValue().getString().length() > 2)
                            dbId = cell.getRichStringCellValue().getString(); // Add the KEGG value in case HMDB is not reported

                        if (dbId.length() > 2)
                            newRow.createCell(0).setCellValue(dbId);                //TODO, KEGG not being added
                    }

                    if (metabolonColumnNumber == 12) { // HMDB
                        if (cell.getRichStringCellValue().getString() != null && cell.getRichStringCellValue().getString().length() > 2)
                            dbId = cell.getRichStringCellValue().getString();

                        if (dbId.length() > 2)
                            newRow.createCell(0).setCellValue(dbId); // "database_identifier"
                    }

                    if (metabolonColumnNumber == 1) //Compound name
                        newRow.createCell(5).setCellValue(cell.getRichStringCellValue().getString()); // "metabolite_identification"

                    if (metabolonColumnNumber >= 13) {  // The sample  concentration values starts at column 13
                        Double cellValue = cell.getNumericCellValue();
                        if (cellValue != null)
                            newRow.createCell((getLastCellNumber()+ metabolonColumnNumber)-13 ).setCellValue(cellValue); // Sample data
                        //TODO, fix the -13 hack!
                    }


                });

                //Adding new value to cell
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
