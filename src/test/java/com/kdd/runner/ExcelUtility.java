package com.kdd.runner;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
//import org.junit.Test;

public class ExcelUtility {

	private static HSSFWorkbook workbook;
	
	public static String testDataPath = "C://DataSheets//TestDataSheetNewPolicy.xls";
	
	private static org.apache.poi.ss.usermodel.Cell Cell;
	private static DataFormatter fmt = new DataFormatter();
	private static String sheetName="Roles";
	public static String testCaseID="TC003";
	//String baseDirectory = System.getProperty("user.dir");
	
	public static void main(String[] args) throws IOException {
		
		List<Map<String,String>> testDataInMap=getMapTestData(sheetName,testCaseID);
		int mapsize=testDataInMap.size();
		System.out.println("Map Size :"+mapsize);
		System.out.println("Data1 :"+testDataInMap.get(0).get("Data10"));
		
	}
		

	public static List<Map<String,String>> getMapTestData(String sheetName,String testCaseID) throws IOException
	  {
		List<Map<String,String>>testDataAllRows=null;
		
		Map<String,String> testData=null;
		
		try 
		{
			System.out.println("Excel Path inside setExcelFile :" + testDataPath);

			FileInputStream excelFile = new FileInputStream(testDataPath);

			workbook = new HSSFWorkbook(excelFile);
			HSSFSheet sheet = workbook.getSheet(sheetName);
			
			int lastRowNumber=sheet.getLastRowNum()+1;
			int lastColNumber=sheet.getRow(0).getLastCellNum();
			
		    // test Main For loop for List Based on Test Case Name
			int TCEndCounterRow=0;
			int startpoint=0;
			for( int count=1;count<lastRowNumber;count++)
			{
				HSSFSheet excelSheet = workbook.getSheet(sheetName);
				
				Cell cellTest=excelSheet.getRow(count).getCell(0);
				String RowTestNameCellData=cellTest.getStringCellValue().toString();
				if(testCaseID.equalsIgnoreCase(RowTestNameCellData))
				   {
				    //
					System.out.println("TestCase Found At Row :"+count);
						
					startpoint=startpoint+1;
					System.out.println("startpoint :"+startpoint);
					TCEndCounterRow=count;
					//
			      } // main if end
			} // Main For loop end	

	  	
	  	int TCstartPointer=(TCEndCounterRow-startpoint)+1;//plus 1
	  	int TCEndPointer=(TCstartPointer+startpoint)-1;//minus 1
	  	
	  	System.out.println("startpointStart :"+TCstartPointer);
	  	System.out.println("finalcount :"+TCEndPointer);
	  	
			List list=new ArrayList();
			
			for(int i=0;i<lastColNumber;i++)
			{
				
				HSSFRow row=sheet.getRow(0);
				HSSFCell cell=row.getCell(i);
				String rowHeader=cell.getStringCellValue().trim();
				System.out.println("rowHeader :"+rowHeader);
				list.add(rowHeader);
				
			}
			
			testDataAllRows=new ArrayList<Map<String,String>>();
			
			//for(int j=1;j<lastRowNumber;j++)
			for(int j=TCstartPointer;j<TCEndPointer+1;j++)
			{
				HSSFRow row=sheet.getRow(j);
				
				testData=new TreeMap<String,String>(String.CASE_INSENSITIVE_ORDER);
				
			  for(int k=1;k<lastColNumber;k++)
				{
					HSSFCell cell=row.getCell(k);
					String colValue=cell.getStringCellValue().trim();
					System.out.println("colValue :"+colValue);
										
					testData.put((String) list.get(k), colValue);
				}
				testDataAllRows.add(testData);
			}
					
		}
		catch(FileNotFoundException e)
		{
			e.printStackTrace();
		}
		
		return testDataAllRows;
	 }
	
		
}
