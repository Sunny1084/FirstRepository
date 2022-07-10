package JXLAssignment;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ReadWrite {
	
	public static void readWriteData() throws BiffException, IOException, RowsExceededException, WriteException {
		
		File file = new File("../JavaClass/Files/TestData.xls");
		Workbook wb = Workbook.getWorkbook(file);
		Sheet ws = wb.getSheet(0);
		int rows = ws.getRows();
		int cols = ws.getColumns();

		File writeFile = new File("../JavaClass/Files/TestWrite.xls");
		
		WritableWorkbook wk = Workbook.createWorkbook(writeFile);
		WritableSheet sheet = wk.createSheet("Data", 0);
		
		for (int i = 0; i < rows; i++ ) {
			for (int j = 0; j < cols; j++) {
				Cell value = ws.getCell(j, i);
				String data = value.getContents();
				Label label = new Label(j, i, data);
				sheet.addCell(label);
			}
		}
		wk.write();
		wk.close();
	}

	public static void main(String[] args) throws RowsExceededException, BiffException, WriteException, IOException {

		readWriteData();
	}

}
