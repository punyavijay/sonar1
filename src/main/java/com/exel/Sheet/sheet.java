package com.exel.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.security.Key;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.JsonArray;
import com.google.gson.JsonObject;


public class sheet
	{
		private static final String name ="/home/developers/vj/file.xlsx";
		
		
		public static void main(String[] args) throws FileNotFoundException 
		{
			JsonObject obj=new JsonObject();
			HashMap<Integer,String> map=new HashMap<Integer,String>();
			ArrayList list=new ArrayList();
			
			try 

			{
				FileInputStream file=new FileInputStream(new File(name));
				Workbook workbook=new XSSFWorkbook(file);
				DataFormatter dataformatter=new DataFormatter();
				Iterator<Sheet> sheets=workbook.sheetIterator();
				
				
				while(sheets.hasNext())
				{
					
					Sheet sh=sheets.next();
					String SheetName=sh.getSheetName();
					Iterator<Row> rowIterator=sh.iterator();
					
					
					while(rowIterator.hasNext() )
					{
						Row row=rowIterator.next();
						Iterator<Cell> cellIterator=row.iterator();
						int key=1;
						while(cellIterator.hasNext())
						{
							Cell cell=cellIterator.next();
							String cellValue=dataformatter.formatCellValue(cell);
							
							if(row.getRowNum()==0 && !(cellValue.isEmpty()))
							{
								map.put(key, cellValue);
								
							}
							if(row.getRowNum()!=0)
							{
								map.put(key, cellValue);
							}
							key++;
						}
						
//						for(Map.Entry m1 : map.entrySet())
//						{    
//							System.out.print(m1.getKey()+" "+m1.getValue()+" \t\t");  
//							
//						}
//						System.out.println();
						
						
						if(SheetName.equals("Sheet1"))
						{

							System.out.println(map);
						}
						
					}
					
					
				}
				
				workbook.close();
			} 
			catch (Exception e) 
			{
				e.printStackTrace();
			}
		}
	}