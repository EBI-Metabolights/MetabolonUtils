package uk.ac.ebi.metabolights.utils.metabolonutils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.isatools.isacreator.configuration.io.ConfigXMLParser;
import org.isatools.isacreator.spreadsheet.model.TableReferenceObject;
import org.isatools.isatab.configurator.schema.IsaTabConfigurationType;
import org.isatools.isatab.configurator.schema.IsatabConfigFileDocument;
import org.isatools.plugins.metabolights.assignments.model.Metabolite;

import java.io.*;
import java.util.ArrayList;
import java.util.Vector;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
import static org.apache.poi.ss.usermodel.CellType.STRING;

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
    private static int metaboliteCellPos = 1;

    private SearchUtils searchUtils = new SearchUtils();

    public static int getLastCellNumber() {
        return lastCellNumber;
    }

    public static void setLastCellNumber(int lastCellNumber) {
        FileUtils.lastCellNumber = lastCellNumber;
    }

    public TableReferenceObject getMSConfig() {
        return getConfiguration();
    }

    public void convertExcelFile(String fileName) throws IOException, InvalidFormatException {
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        System.out.println("Creating a cloned sheet for annotations, renaming the existing sheet to '"+metabolonSheetName+"'");
        Workbook workbook = WorkbookFactory.create(new File(fileName));
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        System.out.println("The original sheet has "+sheet.getPhysicalNumberOfRows() +" number of rows and "+row.getLastCellNum()+" number of columns. This results in a matrix of "+ sheet.getPhysicalNumberOfRows()*row.getLastCellNum() +" cells");
        workbook.setSheetName(workbook.getSheetIndex(sheet), metabolonSheetName);
        workbook.cloneSheet(0); //Copy the existing sheet at position 0

        //Retrieving the cloned worksheet
        Sheet clonedMetabolonSheet = workbook.getSheetAt(1);   //Sheet 0 is the original data from Metabolon, so 1 is the cloned sheet
        System.out.println("Renaming the cloned sheet to '"+annotatedSheetName+"'");
        workbook.setSheetName(workbook.getSheetIndex(clonedMetabolonSheet), annotatedSheetName);

        //Annotate each row with the row type
        System.out.println("Annotate each row with the row type");
        annotateMetabolonSheet(clonedMetabolonSheet);

        //Duplicate and split all rows that has "/" in the compound name
        System.out.println("Duplicate and split all rows that has '/' in the compound name");
        duplicateMetabolonRows(clonedMetabolonSheet);

        // Create the new sheet for MAF
        System.out.println("Create the new sheet for MAF");
        Sheet newSheet = addStandardHeaderRow(workbook);
        newSheet = addSampleHeaderRow(newSheet, clonedMetabolonSheet);
        System.out.println("Add Metabolon data to the MAF sheet");
        addMetabolonData(newSheet, clonedMetabolonSheet);

        // Write the output to a new Excel file
        FileOutputStream fileOut = new FileOutputStream("MetabolonPeakAreaTable_MAF.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    private void annotateMetabolonSheet(Sheet clonedMetabolonSheet){
        //Annotation rows from the original Metabolon sheet
        clonedMetabolonSheet.forEach(row -> {
            int rowNum = row.getRowNum();
            System.out.println("Annotating row: "+rowNum);
            annotateMetabolonData(row, rowNum);
        });
    }

    private void duplicateMetabolonRows(Sheet annotatedMetabolonSheet){
        //Duplicate rows in the annoted Metabolon sheet
        ArrayList<Integer> dupRows = rowsToDuplicate(annotatedMetabolonSheet);

        AtomicInteger additionalRowsToMove = new AtomicInteger();
        dupRows.forEach(row -> {
            System.out.println("Duplicating row "+row);
            copyRows(annotatedMetabolonSheet, row+ additionalRowsToMove.getAndIncrement());
        });
    }

    protected void copyRows(Sheet worksheet, int rowNum) {
        Row sourceRow = worksheet.getRow(rowNum);
        worksheet.shiftRows(rowNum, worksheet.getLastRowNum(), 1);  //Now the source row is at rowNum+1
        Row newRow = worksheet.createRow(rowNum);

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            Cell oldCell = sourceRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); //The cell we are reading the data from. If the cell us null, return a blank value
            Cell newCell = newRow.createCell(i); //New empty cell to put the data into

            // Copy style from old cell and apply to new cell
            CellStyle cellStyle = oldCell.getCellStyle();
            if (cellStyle != null) {
                CellStyle newCellStyle = worksheet.getWorkbook().createCellStyle();
                newCellStyle.cloneStyleFrom(cellStyle);
                newCell.setCellStyle(newCellStyle);
            }

            // Set the cell data value
            switch (oldCell.getCellTypeEnum()) {
                case BLANK: break;
                case BOOLEAN: newCell.setCellValue(oldCell.getBooleanCellValue()); break;
                case ERROR: newCell.setCellErrorValue(oldCell.getErrorCellValue()); break;
                case NUMERIC: newCell.setCellValue(oldCell.getNumericCellValue()); break;
                case STRING:
                    String metabolite = sourceRow.getCell(i).getRichStringCellValue().toString();
                    newCell.setCellValue(metabolite);   //First, set the new value
                    if (metabolite.contains("/") && i == 1) {  // The metabolite name is always in position 1 (2nd column in the sheet)
                        String[] metabolites = extractMetabolites(metabolite);
                        if (metabolites.length == 2) {
                            newCell.setCellValue(metabolites[0]); //1st metabolite name
                            oldCell.setCellValue(metabolites[1]); //2ns metabolite name
                        }
                    } else {
                        newCell.setCellValue(sourceRow.getCell(i).getRichStringCellValue());
                    }
                    break;
                default: break;
            }
        }
    }

    private String[] extractMetabolites(String metabolite){
        String[] metabolites = null;
        int slashPos = 0, startPos = 0, endPos = 0;

        if (metabolite.contains("/"))
            slashPos = metabolite.indexOf("/");

        if (metabolite.contains("(")) {
            startPos = metabolite.indexOf("(");
            endPos = metabolite.indexOf(")");
        }


        if (startPos == 0) // No lipids reported, safe to split using "/"
            return metabolite.split("/");

        //Ignore this pattern "( / )" as this is a lipid
        //Do not split lipids if they use slash in the naming
        //replace "/" with "#" so we can split on # only
        if (slashPos < startPos && slashPos < endPos){ // Lipid in position 2    ( "compound/lipid(1:2/2:1)" )
            metabolite = metabolite.replace("/","#");

        } else if (slashPos > startPos && slashPos > endPos){ // Lipid in position 1 ( "lipid(1:2/2:1)/compound" )
            metabolite = metabolite.replace(")/",")#");
        }

        return metabolite.split("#");
    }


    private ArrayList<Integer> rowsToDuplicate(Sheet annoatedSheet){
        ArrayList<Integer> dupRows = new ArrayList<>();

        annoatedSheet.forEach( row -> {
            row.forEach(cell -> {
                if (cell.getColumnIndex() == metaboliteCellPos ) {
                    String textValue = cell.getRichStringCellValue().getString();

                    if (textValue.contains("/") && !textValue.contains("(")){
                        dupRows.add(cell.getRowIndex()); //Simple split as there are no brackets
                        System.out.println(" - Compound at row "+ row.getRowNum() +" needs splitting: "+ textValue);
                    }  else if (textValue.contains("/") && textValue.contains("(")) {  //Ok, a more complicated split

                        //Let's map out where the brackets and slashes are
                        String charPos ="";
                        for(int i = 0, n = textValue.length() ; i < n ; i++) {
                            switch (textValue.charAt(i)) {
                                case '/':  charPos = charPos + "/"; break;
                                case '(':  charPos = charPos + "("; break;
                                case ')':  charPos = charPos + ")"; break;
                            }
                        }

                        String noEmptyBrackets = charPos.replaceAll("\\(\\)","");
                        if (noEmptyBrackets.equals("/")) {  //After removing empty brackets (), we are left with slashes that divide two compounds
                            dupRows.add(cell.getRowIndex());
                            System.out.println(" - Compound at row "+ row.getRowNum() +" needs splitting: "+ textValue);
                        } else {
                            System.out.println(" - Compound at row "+ row.getRowNum() +" should not be split: "+ textValue);
                        }

                    }
                }
            });
        });

        return dupRows;
    }

    private Sheet addStandardHeaderRow(Workbook workbook){
        Sheet newSheet = workbook.createSheet(MAFSheetName);
        Row headerRow = newSheet.createRow(0);
        TableReferenceObject mafTable = getMSConfig();  //Header values from the config file

        Vector<String> standardHeaders =  mafTable.getHeaders(); //Get all the headers from the config file
        for (int i = 0; i < standardHeaders.size(); i++) {
            if (i>0)  //Skip the first row as this only has the row-number ("Row No.")
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
                    String cellValue = getNumericOrDoubleCellValue(cell);
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

    private String getNumericOrDoubleCellValue(Cell cell){
        String cellValue = "";

        //Sample names can be text and/or numbers
        if (cell.getCellTypeEnum() == STRING){
            if (cell.getRichStringCellValue().getString() != null)
                cellValue = cell.getRichStringCellValue().getString();
        } else if(cell.getCellTypeEnum() == NUMERIC) {
            Double numCellValue = cell.getNumericCellValue();
            if (numCellValue != null)
                cellValue = numCellValue.toString();

        }

        if (cellValue.equals("."))
            return "";

        return cellValue;
    }

    /**
     * Create a new new empty row in the sheet
     * @param newSheet
     * @return Row, with empty values
     */
    private Row getNewRow(Sheet newSheet){
        int numberOfCells = (int) newSheet.getRow(0).getLastCellNum();  // Get the header row
        int newRowNum = newSheet.getLastRowNum();;
        Row newRow = newSheet.createRow(newRowNum+1); //Don't overwrite the last or header row.
        for (int i = 0; i < numberOfCells; i++) {
            newRow.createCell(i).setCellValue("");   //Add all the empty cells first
        }

        return newRow;
    }

    private void addMetabolonData(Sheet newSheet, Sheet metabolonSheet){

        //Output the header row first
        //newSheet.getRow(0).forEach(newCell -> {
        //    printCellValue(newCell);
        //});
        //System.out.println(); //Only to ensure line breaks when printing to screen

        metabolonSheet.forEach(row -> { //Loop through the Metabolon sheet
            int rowNum = row.getRowNum();
            System.out.println("Adding Metabolon data for row: "+rowNum);

            String dataRowType = row.getCell(0).getRichStringCellValue().toString();

            //Retrieving cells
            if (dataRowType.equals(dataAnnotation)){

                Row newRow = getNewRow(newSheet);

                //Lamda requirement
                final String[] dbId = { null };
                final String[] metName = { null };

                row.forEach(cell -> {
                    int metabolonColumnNumber = cell.getColumnIndex();

                    if (metabolonColumnNumber == metaboliteCellPos) { //Compound name
                        metName[0] = cell.getRichStringCellValue().getString();
                        newRow.createCell(4).setCellValue(metName[0]); // "metabolite_identification"
                    }

                    if (metabolonColumnNumber == 11 || metabolonColumnNumber == 12) { // KEGG (11) or HMDB (12)
                        if (cell.getRichStringCellValue().getString() != null && cell.getRichStringCellValue().getString().length() > 2)
                            dbId[0] = cell.getRichStringCellValue().getString(); // to be used for "database_identifier"
                        // Adding KEGG first, in case HMDB is not reported
                    }

                    if (dbId[0] != null && dbId[0].length() > 2)
                        newRow.createCell(0).setCellValue(dbId[0]); // "database_identifier"


                    if (metabolonColumnNumber == 8) { //Mass
                        Double cellValue = cell.getNumericCellValue();
                        if (cellValue != null)
                            newRow.createCell(5).setCellValue(cell.getNumericCellValue()); // "mass_to_charge"
                    }

                    if (metabolonColumnNumber >= 13) {  //The sample concentration values starts at column 13
                        String cellValue = getNumericOrDoubleCellValue(cell);
                        Double doubleCellValue = 0.00;
                        if (cellValue != null) {
                            if (cellValue != "")
                                doubleCellValue = Double.parseDouble(cellValue);

                            newRow.createCell((getLastCellNumber() + metabolonColumnNumber) - 13).setCellValue(doubleCellValue); // Sample data
                            //TODO, fix the -13 hack!
                        }
                    }

                });

                if (metName[0] != null) {

                    Metabolite met;
                    String cleanMetName = metName[0].replaceAll("\\*",""); //Get rid of "*" (astrix) in compound names before searching
                    met = searchUtils.getMetaboliteInformation(dbId[0], cleanMetName);

                    if (met.getIdentifier() == null)
                        met = searchUtils.getMetaboliteInformation(null, cleanMetName); //The compounds may be synonyms, so try ChEBI until we (may) find one

                    if (met != null) { // Add and/or replace with MetaboLights WS search results

                        if (met.getIdentifier() != null)
                            newRow.createCell(0).setCellValue(met.getIdentifier());

                        if (met.getFormula() != null)
                            newRow.createCell(1).setCellValue(met.getFormula());

                        if (met.getSmiles() != null)
                            newRow.createCell(2).setCellValue(met.getSmiles());

                        if (met.getInchi() != null)
                            newRow.createCell(3).setCellValue(met.getInchi());

                    }
                }

                //Printing new value in cell
                //if (newRow != null)
                //    newRow.forEach(newCell -> {
                //        printCellValue(newCell);
                //    });
                //System.out.println();

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

    private TableReferenceObject getConfiguration(){
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

            if (parser.getTables().size() > 0)
                return parser.getTables().get(0);

        } catch (XmlException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return tableReferenceObject;
    }

}
