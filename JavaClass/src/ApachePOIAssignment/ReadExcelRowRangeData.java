package ApachePOIAssignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelRowRangeData {

	public static void readRangeRow(int initialRow, int endRow) throws IOException {
		File file = new File("../JavaClass/Files/LoginData.xlsx");
		FileInputStream inputFile = new FileInputStream(file);
		XSSFWorkbook workBook = new XSSFWorkbook(inputFile);
		XSSFSheet sheet = workBook.getSheetAt(0);
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0 ; i < rows ; i ++ ) {
			XSSFRow row = sheet.getRow(i);
			int cells =  row.getPhysicalNumberOfCells();
			for (int j = 0 ; j < cells ; j++) {
				XSSFCell cell = row.getCell(j);
				if (i >= initialRow && i <= endRow) {
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

		readRangeRow(0, 1);
	}

}
