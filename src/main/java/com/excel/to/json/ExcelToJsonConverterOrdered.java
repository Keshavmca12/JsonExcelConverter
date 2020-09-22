package com.excel.to.json;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.databind.ObjectMapper;

public class ExcelToJsonConverterOrdered {

	public static void main(String[] args) {
		Map<String, Object> data = getExcelData();
		System.out.println(data);
		try {
			new ObjectMapper().writerWithDefaultPrettyPrinter().writeValue(new File("enFromExcel.json"), data);
			/*
			 * String json = new
			 * ObjectMapper().writerWithDefaultPrettyPrinter().writeValueAsString(data);
			 * writeJsonToFile(json);
			 */
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/*
	 * public static void writeJsonToFile(String json) { try (FileWriter file = new
	 * FileWriter("enFromExcel.json")) { file.write(json); file.flush(); } catch
	 * (IOException e) { e.printStackTrace(); } }
	 */

	@SuppressWarnings({ "deprecation", "unchecked" })
	public static Map<String, Object> getExcelData() {
		Map<String, Object> resultData = new LinkedHashMap<>();
		FileInputStream file = null;

		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = null;

		try {
			file = new FileInputStream(new File("jsonToExcel.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				String cellKey = null;

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						// System.out.print(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_STRING:
						String cellData = cell.getStringCellValue();
						// System.out.print(cellData + " s");
						if (null == cellKey) {
							cellKey = cell.getStringCellValue();
						} else {
							String[] keyList = cellKey.split("[.]");
							Map<String, Object> dataMap = new LinkedHashMap<>();
							switch (keyList.length) {
							case 1:
								resultData.put(keyList[0], cellData);
								break;
							case 2:
								if (!resultData.containsKey(keyList[0])) {
									resultData.put(keyList[0], dataMap);
								} else {
									dataMap = (Map<String, Object>) resultData.get(keyList[0]);
								}
								dataMap.put(keyList[1], cellData);
								resultData.put(keyList[0], dataMap);
								break;
							case 3:

								if (!resultData.containsKey(keyList[0])) {
									resultData.put(keyList[0], dataMap);
								} else {
									dataMap = (Map<String, Object>) resultData.get(keyList[0]);
								}

								Map<String, Object> dataMap1 = new LinkedHashMap<>();
								if (dataMap.get(keyList[1]) == null) {
									dataMap.put(keyList[1], dataMap1);
								} else {
									dataMap1 = (Map<String, Object>) dataMap.get(keyList[1]);
								}
								dataMap1.put(keyList[2], cellData);
								dataMap.put(keyList[1], dataMap1);
								resultData.put(keyList[0], dataMap);
								break;

							}
							cellKey = null;
						}
						break;
					}
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				workbook.close();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			try {
				file.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}
		return resultData;
	}

}
