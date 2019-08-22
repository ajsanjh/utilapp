import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Iterator;

/**
 * Created by rajeevkumarsingh on 18/12/17.
 */

public class PrintSheet {

    private static String fileP;

	public static void main(String[] args) throws IOException, InvalidFormatException {
    	
		BufferedReader reader =  
				new BufferedReader(new InputStreamReader(System.in));
		System.out.println("File path: ");
		fileP = reader.readLine(); 
		System.out.println("Working... \n\n");

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(fileP));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets... ");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        workbook.close();
    }

}