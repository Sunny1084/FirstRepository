package ApachePOIAssignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReadCellData {
	
	public static void readCellData(int rowNo, int colNo) throws IOException {
		File file = new File("../JavaClass/Files/LoginData.xlsx");
		FileInputStream inputFile = new FileInputStream(file);
		XSSFWorkbook workBook = new XSSFWorkbook(inputFile);
		XSSFSheet sheet = workBook.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		System.out.println(rows);
		for (int i = 0 ; i < rows ; i ++ ) {
			XSSFRow row = sheet.getRow(i);
			int cells =  row.getPhysicalNumberOfCells();
			for (int j = 0 ; j < cells ; j++) {
				XSSFCell cell = row.getCell(j);
				if (i == rowNo && j == colNo) {
					switch(cell.getCellType()) {
					case NUMERIC:
						System.out.println(cell.getNumericCellValue());
						break;
					case STRING:
						System.out.println(cell.getStringCellValue());
						break;
					}
				}
			}
		}
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub


		readCellData(1, 0);
		
		
	}

}
