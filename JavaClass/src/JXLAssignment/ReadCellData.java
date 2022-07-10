package JXLAssignment;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ReadCellData {
	
	public static void ReadDataBasedUponRowNoandColNo(int rowNo, int colNo) throws BiffException, IOException {
		
		File file = new File("../JavaClass/Files/TestData.xls");
		Workbook wb = Workbook.getWorkbook(file);
		Sheet ws = wb.getSheet(0);
		int rows = ws.getRows();
		int cols = ws.getColumns();
		
		for (int i = 0; i < rows; i++ ) {
			for (int j = 0; j < cols; j++) {
				Cell value = ws.getCell(j, i);
				if (i == rowNo && j == colNo) {
				System.out.println(value.getContents());
				break;
				}
			}
		}
	}
	

	public static void main(String[] args) throws BiffException, IOException {
		// TODO Auto-generated method stub

		ReadDataBasedUponRowNoandColNo(2, 1);
	}

}
