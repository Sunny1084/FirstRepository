package ApachePOIAssignment;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFIleWrite {

	public static void main(String[] args) throws IOException {

		File file = new File("../JavaClass/Files/WriteData.xlsx");
		FileOutputStream outputFile = new FileOutputStream(file);
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet("Output");
		 for (int i = 0 ; i < 3 ; i++) {
			 XSSFRow rows = sheet.createRow(i);
			 for (int j = 0 ; j < 3 ; j++ ) {
				 XSSFCell cell = rows.createCell(j);
				 cell.setCellValue("Sunny");
			 }
		 }
		 workBook.write(outputFile);
		 outputFile.flush();
		 outputFile.close();
	}

}
