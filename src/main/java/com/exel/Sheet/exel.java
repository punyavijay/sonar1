package com.exel.Sheet;


import java.io.BufferedReader;
import java.io.*;
import java.io.FileInputStream;
import java.io.FileReader;
import java.util.Scanner;
import java.io.IOException;
import java.io.InputStreamReader;

public class exel 
{
	public static void main(String[] args) throws IOException 
	{
		
			FileInputStream fstream = new FileInputStream("/home/developers/vj/first.xls");
			BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
			
				String strLine;
				while((strLine = br.readLine()) != null) 
				{
					System.out.println(strLine);
				}
				br.close();
			} 
			
				
				
			
		
	
	
	
}
