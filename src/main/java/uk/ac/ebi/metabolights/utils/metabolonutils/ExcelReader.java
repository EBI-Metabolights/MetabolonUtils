
package uk.ac.ebi.metabolights.utils.metabolonutils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;

/**
 * Created by rajeevkumarsingh on 18/12/17.
 */

public class ExcelReader {

    public static final String ExcelFile = ExcelReader.class.getClassLoader().getResource("." + File.separator + "MetabolonPeakAreaTable.xlsx").getFile();

    public static void main(String[] args) {

        FileUtils fileUtils = new FileUtils();

        try {
            fileUtils.convertExcelFile(ExcelFile);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

}


