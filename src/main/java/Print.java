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

public class Print {

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
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 3. Or you can use Java 8 forEach loop with lambda
//        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
//        sheet.forEach(row -> {
//            row.forEach(cell -> {
//                printCellValue(cell);
//            });
//            System.out.println();
//        });

        // Closing the workbook
        workbook.close();
    }

    private static void printCellValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                System.out.print(cell.getCellFormula());
                break;
            case BLANK:
                System.out.print("");
                break;
            default:
                System.out.print("");
        }

        System.out.print("\t");
    }
}