package test;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelProgram {
	FileInputStream f;
	XSSFWorkbook x;
	XSSFSheet s;

	public void getExcel() throws IOException {
		f = new FileInputStream("E:\\java\\java notes\\Book1.xlsx");
		x = new XSSFWorkbook(f);
		s = x.getSheet("Sheet1");
	}

	public void getData(int i, int j) throws IOException {
		Row r = s.getRow(i);
		Cell c = r.getCell(j);
		switch (c.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			System.out.println(c.getStringCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			System.out.println(c.getNumericCellValue());
			break;

		}
	}
}
