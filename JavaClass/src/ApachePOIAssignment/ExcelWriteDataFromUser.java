package ApachePOIAssignment;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteDataFromUser {

	public static void writeDataUser() throws IOException {
		
		File file = new File("../JavaClass/Files/WriteData.xlsx");
		FileOutputStream outputFile = new FileOutputStream(file);
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("Output");
		Scanner sc = new Scanner(System.in);
		String value = sc.next();
		for (int i = 0 ; i < 3 ; i++) {
			XSSFRow rows = sheet.createRow(i);
			for (int j = 0 ; j < 3 ; j++ ) {
				 XSSFCell cells = rows.createCell(j);
				 cells.setCellValue(value);
			}
		}
		workBook.write(outputFile);
		outputFile.flush();
		outputFile.close();
		sc.close();
	}
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		writeDataUser();
	}

}
