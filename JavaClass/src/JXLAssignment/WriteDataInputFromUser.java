package JXLAssignment;

import java.io.File;
import java.io.IOException;
import java.util.Scanner;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteDataInputFromUser {

	public static void writeData(int rowCount, int colCount) throws IOException, RowsExceededException, WriteException {

		File file = new File("../JavaClass/Files/TestWrite.xls");
		
		WritableWorkbook wk = Workbook.createWorkbook(file);
		WritableSheet sheet = wk.createSheet("Data", 0);
		
		Scanner sc = new Scanner(System.in);
		String input = sc.next();
		for (int i = 0; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				Label label = new Label(j, i, input);
				sheet.addCell(label);
			}
		}
		wk.write();
		wk.close();
		sc.close();
	}
	public static void main(String[] args) throws RowsExceededException, WriteException, IOException {
		// TODO Auto-generated method stub

		writeData(2,3);
	}

}
