package ApachePOIAssignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileRead {

	public static void main(String[] args) throws IOException {
		
		File file = new File("../JavaClass/Files/LoginData.xlsx");
		
		FileInputStream inputFile = new FileInputStream(file);
		
		XSSFWorkbook workBook = new XSSFWorkbook(inputFile);
		
		XSSFSheet sheet = workBook.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		
		for (int i = 0; i < rows; i++) {
			XSSFRow rowValue = sheet.getRow(i);
			int columns = rowValue.getPhysicalNumberOfCells();
			for (int j = 0 ; j < columns ; j++) {
				XSSFCell colValue = rowValue.getCell(j);
				switch (colValue.getCellType()) {
				case NUMERIC:
					System.out.println(colValue.getNumericCellValue());
					break;
				case STRING:
					System.out.println(colValue.getStringCellValue());
					break;
				}
			}
		}
		inputFile.close();
	}
}
