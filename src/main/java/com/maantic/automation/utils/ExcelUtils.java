package com.maantic.automation.utils;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class ExcelUtils {
	private ExcelUtils() {

	}

	public static List<Map<String, String>> getExcelData(String sheetName) {
		List<Map<String, String>> list = null;
		// copy files
		File sourceExcel = new File(Constants.TEST_DATA_SHEET_PATH);
		File dstExcel = new File(Constants.TEST_OUT_DATA_SHEET_PATH);
		try {
			FileUtils.copyFile(sourceExcel, dstExcel);
		} catch (IOException e) {
			e.printStackTrace();
		}

		FileInputStream fs;

		try {
			// System.out.println("Data File"+Constants.TEST_DATA_SHEET_PATH);
			fs = new FileInputStream(Constants.TEST_DATA_SHEET_PATH);
			XSSFWorkbook wb = new XSSFWorkbook(fs);
			XSSFSheet wSheet = wb.getSheet(sheetName);

			int lastRowNum = wSheet.getLastRowNum();
			int lastColNum = wSheet.getRow(0).getLastCellNum();

			Map<String, String> dataMap = null;
			list = new ArrayList<>();

			for (int i = 1; i <= lastRowNum; i++) {
				dataMap = new HashMap<>();
				String value = "";
				for (int k = 0; k < lastColNum; k++) {
					String key = wSheet.getRow(0).getCell(k).getStringCellValue();
					if (wSheet.getRow(i).getCell(k) == null) // getCellType() == CellType.BLANK ||
																// wSheet.getRow(i).getCell(k).getStringCellValue().trim().isEmpty())
						value = "NA";
					else
						value = wSheet.getRow(i).getCell(k).getStringCellValue();
					dataMap.put(key, value);
				}
				list.add(dataMap);
			}
			return list;
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}

	}

	public static void writeExcelData(String writeOutput, String ruleType, int colNum) {
		XSSFWorkbook workbook = null;
		try {
			FileInputStream file = new FileInputStream(new File(Constants.TEST_DATA_SHEET_PATH));
			workbook = new XSSFWorkbook(file);
			XSSFSheet wSheet = workbook.getSheet(Constants.EXCEL_SHEET_NAME);

			int lastRowNum = wSheet.getLastRowNum();
			// int lastColNum = wSheet.getRow(0).getLastCellNum();

			for (int i = 1; i <= lastRowNum; i++) {
				// check if current row's ruletype is same & pass/fail is NULL
				if (wSheet.getRow(i).getCell(0).toString().equals(ruleType)
						&& wSheet.getRow(i).getCell(18).toString().equals("")) {
					XSSFCell cell = wSheet.getRow(i).createCell(colNum);
					// XSSFCell cell = wSheet.getRow(i).getCell(colNum);
					cell.setCellType(CellType.STRING);
					cell.setCellValue(writeOutput);
					file.close();
					break;
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			FileOutputStream out = new FileOutputStream(new File(Constants.TEST_DATA_SHEET_PATH));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("Output generated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void writeBlankExcelData(String writeOutput, int colNum, int colNum2) {// regenerate i/p file for user
		XSSFWorkbook workbook = null;
		try {
			FileInputStream file = new FileInputStream(new File(Constants.TEST_DATA_SHEET_PATH));
			workbook = new XSSFWorkbook(file);
			XSSFSheet wSheet = workbook.getSheet(Constants.EXCEL_SHEET_NAME);

			int lastRowNum = wSheet.getLastRowNum();
			// int lastColNum = wSheet.getRow(0).getLastCellNum();

			for (int i = 1; i <= lastRowNum; i++) {
				// check if current row's ruletype is same & pass/fail is NULL
				// if (wSheet.getRow(i).getCell(0).toString().equals(ruleType)) {// &&
				// wSheet.getRow(i).getCell(18) == null){
				XSSFCell cell = wSheet.getRow(i).createCell(colNum);
				XSSFCell cell2 = wSheet.getRow(i).createCell(colNum2);
				// XSSFCell cell = wSheet.getRow(i).getCell(colNum);
				cell.setCellType(CellType.STRING);
				cell.setCellValue(writeOutput);
				cell2.setCellType(CellType.STRING);
				cell2.setCellValue(writeOutput);
				// file.close();
				// break;
				// }
			}
			file.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			FileOutputStream out = new FileOutputStream(new File(Constants.TEST_DATA_SHEET_PATH));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("Blank i/p sheet generated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void writeOutputFileData(String ruleType) {// to delete rows not needed
		XSSFWorkbook workbook = null;
		try {
			FileInputStream file = new FileInputStream(new File(Constants.TEST_OUT_DATA_SHEET_PATH));
			workbook = new XSSFWorkbook(file);
//			XSSFSheet wSheet = workbook.getSheet(Constants.EXCEL_SHEET_NAME);
			Sheet sheet = workbook.getSheet(Constants.EXCEL_SHEET_NAME);

//			int lastRowNum = wSheet.getLastRowNum();
			// int lastColNum = wSheet.getRow(0).getLastCellNum();

			int rowCount = 1;

			// Iterate through the rows
			for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);

				if (row != null) {
					Cell cell = row.getCell(0);

					// If the cell in column A is "Activity", keep the row
					if (cell != null && ruleType.equals(cell.getStringCellValue())){              //passing args[0] from CMD
//							|| "RuleType".equals(cell.getStringCellValue()))) {
						if (rowIndex != rowCount) {
							// Shift the row up							
							Row newRow = sheet.getRow(rowCount);
                            if (newRow == null) {
                                newRow = sheet.createRow(rowCount);
                            }
                            copyRow(row, newRow);
						}
						rowCount++;
					}
				}
			}

			// Remove the remaining rows at the bottom
			int lastRowNum = sheet.getLastRowNum();
			for (int i = lastRowNum; i >= rowCount; i--) {
				Row rowToDelete = sheet.getRow(i);
				if (rowToDelete != null) {
					sheet.removeRow(rowToDelete);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			FileOutputStream out = new FileOutputStream(new File(Constants.TEST_OUT_DATA_SHEET_PATH));
			workbook.write(out);
			workbook.close();
			out.close();
			System.out.println("Output GBT file generated successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// Helper method to copy content from one row to another
	private static void copyRow(Row sourceRow, Row destinationRow) {
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			Cell oldCell = sourceRow.getCell(i);
			if (oldCell != null) {
				Cell newCell = destinationRow.createCell(i);
				newCell.setCellValue(oldCell.getStringCellValue());
				newCell.setCellStyle(oldCell.getCellStyle());
			}
		}
	}

}
