package ApachePOIAssignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReadWriteData {

	public static void readWriteData () throws IOException {
		
		File readFile = new File("../JavaClass/Files/LoginData.xlsx");
		FileInputStream inputFile = new FileInputStream(readFile);
		XSSFWorkbook inputWork = new XSSFWorkbook(inputFile);
		XSSFSheet inputSheet = inputWork.getSheetAt(0);
		int inputRows = inputSheet.getPhysicalNumberOfRows();
		
		File writeFile = new File("../JavaClass/Files/WriteData.xlsx");
		FileOutputStream outputFile = new FileOutputStream(writeFile);
		XSSFWorkbook outputWork = new XSSFWorkbook();
		XSSFSheet outputSheet = outputWork.createSheet("Output");
		
		for ( int i = 0; i < inputRows ; i++) {
			XSSFRow inputRow =inputSheet.getRow(i);
			 int inputCells = inputRow.getPhysicalNumberOfCells();
			XSSFRow outRows = outputSheet.createRow(i);
			for (int j = 0 ; j < inputCells ; j++ ) {
				 XSSFCell inputCell = inputRow.getCell(j);
				 XSSFCell outCell = outRows.createCell(j);
				 
				 switch(inputCell.getCellType()) {
				 case NUMERIC:
					double numValue = inputCell.getNumericCellValue();
					outCell.setCellValue(numValue);
					break;
				 case STRING:
					 String stringValue = inputCell.getStringCellValue();
					 outCell.setCellValue(stringValue);
					 break;
				 }
			}
			
		}
		outputWork.write(outputFile);
		outputFile.flush();
		outputFile.close();
	}
	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub

		readWriteData();
	}

}
