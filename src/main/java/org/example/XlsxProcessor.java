package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class XlsxProcessor {

	public File createXlsxFile(List<User> users) {
		Workbook workbook = new XSSFWorkbook();

		//Create centered bold style for header with background
		CellStyle centerBoldStyle = createHeaderRowStyle(workbook);

		XSSFSheet sheet1 = (XSSFSheet) workbook.createSheet("Users");
		//Create header row
		Row headerRow = sheet1.createRow(0);
		createCell(headerRow, 0, "First Name", centerBoldStyle);
		createCell(headerRow, 1, "Last Name", centerBoldStyle);
		createCell(headerRow, 2, "Age", centerBoldStyle);
		createCell(headerRow, 3, "Email", centerBoldStyle);

		//Populate users with default font and style
		for (int i = 0; i < users.size(); i++) {
			int rowIndex = i + 1;
			Row row = sheet1.createRow(rowIndex);
			row.createCell(0).setCellValue(users.get(i).getFirstName());
			row.createCell(1).setCellValue(users.get(i).getLastName());
			row.createCell(2).setCellValue(users.get(i).getAge());
			row.createCell(3).setCellValue(users.get(i).getEmail());
		}

		//Set auto size to columns to fit content
		sheet1.autoSizeColumn(0);
		sheet1.autoSizeColumn(1);
		sheet1.autoSizeColumn(2);
		sheet1.autoSizeColumn(3);

		//Write output to users.xlsx file
		File file = new File("users.xlsx");
		try (OutputStream fileOut = new FileOutputStream(file)) {
			workbook.write(fileOut);
			workbook.close();
			return file;
		} catch (IOException e) {
			e.printStackTrace();
			return null;
		}
	}

	public List<User> parseXlsxFile(File file) {
		List<User> users = new ArrayList<>();
		try (Workbook workbook = WorkbookFactory.create(file)) {
			Sheet sheet = workbook.getSheetAt(0); //Also it's possible to get sheet by Name

			for (Row row : sheet) {
				if (row.getRowNum() == 0) {
					continue;
				}
				User user = new User();
				user.setFirstName(getCellValue(row.getCell(0)));
				user.setLastName(getCellValue(row.getCell(1)));
				user.setAge((int) Double.parseDouble(getCellValue(row.getCell(2))));
				user.setEmail(getCellValue(row.getCell(3)));
				users.add(user);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return users;
	}

	private CellStyle createHeaderRowStyle(Workbook workbook) {
		CellStyle centerBoldStyle = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		font.setColor(IndexedColors.WHITE.getIndex());
		centerBoldStyle.setFont(font);
		centerBoldStyle.setFillBackgroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		centerBoldStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		centerBoldStyle.setAlignment(HorizontalAlignment.CENTER);
		return centerBoldStyle;
	}

	private void createCell(Row row, int column, String value, CellStyle style) {
		Cell cell = row.createCell(column);
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}

	private String getCellValue(Cell cell) {
		String value;
		switch (cell.getCellType()) {
			case STRING:
				value = cell.getStringCellValue();
				break;
			case NUMERIC:
				value = String.valueOf(cell.getNumericCellValue());
				break;
			default:
				value = "";
		}
		return value;
	}

}
