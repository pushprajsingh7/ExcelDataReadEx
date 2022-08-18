package com.blucursor.ExcelReadData;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * program to Read data from the given excel file using Apache POI library(or
 * any other utility of your choice), and print the unique email for each KRI
 * ID.
 * 
 * @author ss
 *
 */
public class ExcelKRI {

	@SuppressWarnings({ "incomplete-switch", "resource" })
	public void excel() throws IOException {
		FileInputStream fis = new FileInputStream(new File("C:\\Users\\ss\\Downloads\\Slack Data\\KRIData.xlsx"));
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		FormulaEvaluator formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
		Set<String> setmails = new LinkedHashSet<String>();
		Set<String> setheading = new LinkedHashSet<String>();
		for (Row row : sheet) {
			for (Cell cell : row) {
				switch (formulaEvaluator.evaluateInCell(cell).getCellType()) {
				case STRING:
					String[] emails = cell.getStringCellValue().split(",");
					System.out.println(cell.getStringCellValue());
					for (String mails : emails) {
						if (mails.contains("@")) {
							setmails.add(mails);
							}
						else {
							setheading.add("\t" + mails+ "\t");
						}
					}
					break;
				case NUMERIC:
					System.out.print(cell.getNumericCellValue()+"\t");
					break;
				}
			}
		
		}
//		System.out.println(setheading);
//		System.out.println(setmails);
	}
}
