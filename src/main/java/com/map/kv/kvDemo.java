package com.map.kv;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.security.Key;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonObject;

public class kvDemo {
// private static final String name = "/home/developers/Downloads/file.xlsx";
	
	

	public static void main(String[] args) throws FileNotFoundException {
		// Map<String, Value> studentData = new TreeMap<String, Value>();
		try {
			FileInputStream file = new FileInputStream(new File("/home/developers/vj/file.xlsx"));
			Workbook workbook = new XSSFWorkbook(file);
			DataFormatter dataformatter = new DataFormatter();
			Iterator<Sheet> sheets = workbook.sheetIterator();
			int sheetno = 1;
			while (sheets.hasNext()) {
				Sheet sh = sheets.next();
				Iterator<Row> rowIterator = sh.iterator();
				while (rowIterator.hasNext()) {
					JsonObject object = new JsonObject();
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.iterator();
					{
						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();

							String cellValue = dataformatter.formatCellValue(cell);
							if (row.getRowNum() == 0) {

								switch (cell.getColumnIndex()) {
								case 0:
									key1 = cellValue;
									System.out.println("key1 =" + key1);
									break;

								case 1:
									key2 = cellValue;
									System.out.println("key2 =" + key2);
									break;

								case 2:
									key3 = cellValue;
									System.out.println("key3 =" + key3);
									break;

								case 3:
									key4 = cellValue;
									System.out.println("key4 =" + key4);
									break;

								case 4:
									String key5 = cellValue;
									System.out.println("key5 =" + key5);
									break;
								}
							} else {
								switch (cell.getColumnIndex()) {
								case 0:
									object.addProperty(key1, cellValue);
									break;

								case 1:
									object.addProperty(key2, cellValue);
									break;

								case 2:
									object.addProperty(key3, cellValue);
									break;

								case 3:
									object.addProperty(key4, cellValue);
									break;

								default:
									System.out.println("not added to object");
								}
							}
						}
						if (row.getRowNum() != 0)
							System.out.print(object);

						else {
							System.out.println(sh.getSheetName());
						}
					}
					System.out.println();
				}
			}
			System.out.println();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
