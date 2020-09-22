package com.json.to.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.Stack;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

public class JsonToExcelConverter {
	@SuppressWarnings("rawtypes")
	public static void main(String[] args) {
		try {
			JSONParser parser = new JSONParser();
			JSONObject data = (JSONObject) parser.parse(new FileReader("en.json"));// path to the JSON file.
			System.out.println("data" + data);
			Map map = new LinkedHashMap();
			map = iterateJson(data, map);

		//	System.out.println("map" + map);

		//	WriteExcelFile(map);

		} catch (IOException | ParseException e) {
			e.printStackTrace();
		}

	}

	@SuppressWarnings({ "resource", "rawtypes", "unchecked" })
	public static void WriteExcelFile(Map data) {
		// Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create a blank sheet
		XSSFSheet sheet = workbook.createSheet("JsonToExcel");

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object value = data.get(key);
			int cellnum = 0;
			Object[] objArr = { key, value };
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
			}
		}
		try {
			FileOutputStream out = new FileOutputStream(new File("jsonToExcel.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static Map iterateJson(JSONObject jsonObject, Map map) {
		Stack<String> stack = new Stack<>();
		for (Object o : jsonObject.keySet()) {
			String currentKey = o.toString();
			if (jsonObject.get(currentKey) instanceof JSONObject) {
				stack = new Stack<>();
				stack.push(currentKey);
				iterateJsonHandler((JSONObject) jsonObject.get(currentKey), map, stack);
			} else {
				map.put(currentKey, jsonObject.get(currentKey));
			}

		}

		return map;
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static Map iterateJsonHandler(JSONObject jsonObject, Map map, Stack<String> stack) {
		for (Object o : jsonObject.keySet()) {
			String currentKey = o.toString();
			if (jsonObject.get(currentKey) instanceof JSONObject) {
				stack.push(currentKey);
				iterateJsonHandler((JSONObject) jsonObject.get(currentKey), map, stack);
				if (stack.size() > 1) {
					stack.pop();
				}
			} else {
				String putkey = "";
				for (String item : stack) {
					if (putkey.length() == 0) {
						putkey = item;
					} else {
						putkey = putkey + "." + item;
					}
				}
				if (stack.size() > 0) {
					putkey = putkey + "." + currentKey;
				}
				map.put(putkey, jsonObject.get(currentKey));
			}

		}
		return map;
	}

}
