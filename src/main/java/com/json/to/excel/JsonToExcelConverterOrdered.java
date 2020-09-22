package com.json.to.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.Stack;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jettison.json.JSONException;
import org.codehaus.jettison.json.JSONObject;

public class JsonToExcelConverterOrdered {
	@SuppressWarnings("rawtypes")
	public static void main(String[] args) {
		try {
			Path filePath = Paths.get("en.json");
			String json = new String(Files.readAllBytes(filePath));
			JSONObject data = new JSONObject(json);

			Map map = new LinkedHashMap();
			map = iterateJson(data, map);

			System.out.println("map" + map);

			WriteExcelFile(map);

		} catch (Exception e) {
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
	public static Map iterateJson(JSONObject jsonObject, Map map) throws JSONException {
		Stack<String> stack = new Stack<>();
		Iterator<Object> it = jsonObject.keys();
		while (it.hasNext()) {
			Object o = it.next();
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
	public static Map iterateJsonHandler(JSONObject jsonObject, Map map, Stack<String> stack) throws JSONException {
		Iterator<Object> it = jsonObject.keys();
		while (it.hasNext()) {
			Object o = it.next();
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
