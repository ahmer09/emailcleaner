package com.parlamind.emailcleaner;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import java.util.Scanner;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.concurrent.ForkJoinWorkerThread;

@SpringBootApplication
public class EmailcleanerApplication {

	public static void main(String[] args) throws IOException, InvalidFormatException {
		SpringApplication.run(EmailcleanerApplication.class, args);

		Scanner sc = new Scanner(System.in);

		System.out.println("Enter implicit path of Excel:");
		String path = sc.next();

		final String SAMPLE_XLSX_FILE_PATH = path;

		//Creating a Workbook from an excel file
		Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
		// Retrieving the number of sheets in the Workbook
		System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

		Sheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(0); //assume first row is column header row

		for (int col = 0; col < 2; col++) {
			Cell cell = row.getCell(col);
			if (cell != null) {
				String columnname = cell.getStringCellValue();
				System.out.println(columnname);
			}

		}
	}}
