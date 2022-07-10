package JXLAssignment;

import java.io.File;
import java.io.IOException;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelWrite {

	public static void main(String[] args) throws IOException, RowsExceededException, WriteException {

		File file = new File("../JavaClass/Files/TestWrite.xls");
		
		WritableWorkbook wk = Workbook.createWorkbook(file);
		WritableSheet sheet = wk.createSheet("Data", 0);
		
		for (int i = 0; i < 3; i++) {
			for (int j = 0; j < 3; j++) {
				Label label = new Label(j, i, "Sunny");
				sheet.addCell(label);
			}
		}
		wk.write();
		wk.close();
		
	}

}
