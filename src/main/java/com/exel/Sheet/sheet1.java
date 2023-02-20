package com.exel.Sheet;
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

public class sheet1 
{
		private static final String name = "/home/developers/vj/file.xlsx";
		// static JsonObject object=new JsonObject();
		
		public static void main(String[] args) throws FileNotFoundException 
		{
					// Map<String, Value> studentData = new TreeMap<String, Value>();
					try 
					{
						FileInputStream file = new FileInputStream(new File(name));
						Workbook workbook = new XSSFWorkbook(file);
						DataFormatter dataformatter = new DataFormatter();
						Iterator<Sheet> sheets = workbook.sheetIterator();
						int sheetno=1;
						while (sheets.hasNext()) 
						{
							Sheet sh = sheets.next();
							Iterator<Row> rowIterator = sh.iterator();
							int rowCount = 0;
							while (rowIterator.hasNext()) 
							{
									JsonObject object=new JsonObject();
									Row row = rowIterator.next();
									Iterator<Cell> cellIterator = row.iterator();
									if(rowCount!=0)
									{
									
									int cellElementCount=0;
									
									while (cellIterator.hasNext()) 
									{
										Cell cell = cellIterator.next();
										String cellValue = dataformatter.formatCellValue(cell);
										// System.out.println(cellValue);
										
										switch(sheetno)
										{
										//for first table
										case 1:{
										switch(cellElementCount)
										{
										case 0:
										object.addProperty("First Name",cellValue);
										break;
										case 1:
										object.addProperty("Last Name",cellValue);
										break;
										case 2:
										object.addProperty("Gender Name",cellValue);
										break;
										
										default :
										// System.out.println("Error while inserting key,val in Object for first sheet");
										}
										}
										break;
										//for second table
										case 2:{
										
										switch(cellElementCount)
										{
										case 0:
										object.addProperty("Country",cellValue);
										break;
										case 1:
										object.addProperty("Age",cellValue);
										break;
										case 2:
										object.addProperty("Date",cellValue);
										case 3:
										object.addProperty("Id",cellValue);
										break;
										
										default :
										System.out.println("Error while inserting key,val in Object for second sheet");
										}
										}
										break;
										default :{
										System.out.println("Sheet no not incremented");
										}
										}
										
										// studentData.put("Sr. No.", cellValue);
										// System.out.print(cellValue + "\t\t");
										cellElementCount++;
										}
								}
									if(rowCount!=0)
									System.out.println(object);
									rowCount++;
							}
							System.out.println();
							sheetno++;
							
						}
							workbook.close();
						
					} 
					catch (Exception e) 
					{
						e.printStackTrace();
					}
		}
}