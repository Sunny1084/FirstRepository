package ApachePOIAssignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelRowData {
	
	public static void readRowData(int row) throws IOException {
		
		File file = new File("../JavaClass/Files/LoginData.xlsx");
		FileInputStream inputFile = new FileInputStream(file);
		XSSFWorkbook workBook = new XSSFWorkbook(inputFile);
		XSSFSheet sheet = workBook.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0 ; i < rows ; i++) {
			XSSFRow cell = sheet.getRow(i);
			int cells = cell.getPhysicalNumberOfCells();
			for (int j = 0 ; j < cells ; j++) {
				XSSFCell cellValue = cell.getCell(j);
				if (i == row) {
					switch(cellValue.getCellType()) {
					case NUMERIC:
					     System.out.println(cellValue.getNumericCellValue());
					     break;
					case STRING:
						System.out.println(cellValue.getStringCellValue());
						break;
					}
				}
			}
		}
		
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		readRowData(1);
		
	}

}
