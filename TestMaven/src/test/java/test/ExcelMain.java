package test;

import java.io.IOException;

public class ExcelMain {
	public static void main(String[] args) throws IOException {
		ExcelProgram obj = new ExcelProgram();
		obj.getExcel();
		obj.getData(0, 0);
		obj.getData(0, 1);
		obj.getData(1, 0);
		obj.getData(1, 1);
	}
}
