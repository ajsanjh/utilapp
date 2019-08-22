
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Run {

	public static String fileP;
//	public static String fileP = "./test.xlsx";
	private static String baseSheet = "Sheet1";
	private static int rowNum = 1;

	public static void main(String[] args) throws IOException, InvalidFormatException {

		BufferedReader reader =  
				new BufferedReader(new InputStreamReader(System.in));
		System.out.println("File path: ");
		fileP = reader.readLine(); 
		System.out.println("Row: ");
		rowNum = Integer.parseInt(reader.readLine());
		System.out.println("Working... \n\n");
		
		Workbook workbook = WorkbookFactory.create(new File(fileP));

		myformat(workbook);
		
		
		workbook.close();
		System.out.println("\n\nDone!");
	}

	private static void myformat(Workbook w) {
		Sheet base = w.getSheet(baseSheet);
		DataFormatter dataFormatter = new DataFormatter();
		System.out.printf("**********%n%s%n**********%n%s%n************" ,
				dataFormatter.formatCellValue(base.getRow(rowNum).getCell(0)),
				dataFormatter.formatCellValue(base.getRow(rowNum).getCell(1))
				);

	}


}
