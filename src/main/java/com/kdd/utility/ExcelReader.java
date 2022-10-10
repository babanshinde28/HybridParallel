package com.kdd.utility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class ExcelReader {
	private static HSSFWorkbook excelWorkbook;
	private static final Logger log = Logger.getLogger(ExcelReader.class.getName());
	private static final String RUN_MODE_YES = "YES";

	// Baban Added
	private static org.apache.poi.ss.usermodel.Cell Cell;
	private static DataFormatter fmt = new DataFormatter();

	public synchronized static void setExcelFile(String sheetPath) {
		try {
			
			System.out.println("Excel Path inside setExcelFile :" + sheetPath);

			FileInputStream excelFile = new FileInputStream(sheetPath);

			excelWorkbook = new HSSFWorkbook(excelFile);
		} catch (Exception exp) {
			log.error("Exception occured in setExcelFile: ", exp);
		}
	}

	public static synchronized int getNumberOfRows(String sheetName) {

		HSSFSheet excelSheet = excelWorkbook.getSheet(sheetName);

		// int numberOfRows = excelSheet.getPhysicalNumberOfRows();

		int numberOfRows = excelSheet.getLastRowNum()+1;

		System.out.println("Excel getNumberOfRows() :" + numberOfRows);

		log.debug("Number Of Rows: " + numberOfRows);
		return numberOfRows;
	}

	// New Added to Get No Of Columns
	public static synchronized int getNumberOfCols(String testCaseName, String sheetName) {

		HSSFSheet excelSheet = excelWorkbook.getSheet(sheetName);
		
		HSSFRow Row = null; 
		    			 
		// Making the object of excel row 
		int iRowNumCount = getNumberOfRows(sheetName); 
		int iExpectedTestCaseRowNum=0;
		/// ************* Testing
		for(int iRowNum=1;iRowNum<iRowNumCount+1;iRowNum++)
    	{
    		Cell cell=excelSheet.getRow(iRowNum).getCell(0);
    		String RowTestNameCellData=cell.getStringCellValue().toString();
    		
    		if(RowTestNameCellData.equalsIgnoreCase((testCaseName).trim()))
    		{
    			iExpectedTestCaseRowNum=(int)iRowNum;
    			System.out.println("getNumberOfCols()-TestCase Name "+testCaseName+" Found in DataSheet "+sheetName+" At Row Num "+iExpectedTestCaseRowNum+"  For getCellColData:"+iExpectedTestCaseRowNum);
    			break;
    			
    		}
    	}
		
		///// End Testing
		Row= excelSheet.getRow(iExpectedTestCaseRowNum);
		int numberOfCols = Row.getPhysicalNumberOfCells();//Row.getLastCellNum(); 
		System.out.println("Excel getNumberOfCols() :" + numberOfCols);

		log.debug("Number Of Cols: " + numberOfCols);
		return numberOfCols;
	}
  

public synchronized String getCellData(int rowNumb, int colNumb, String sheetName) throws Exception{ 
	  
	  try{
	  
	  //XSSFSheet excelSheet = excelWorkbook.getSheet(sheetName); 
	  //XSSFCell cell = excelSheet.getRow(rowNumb).getCell(colNumb);
	  //log.debug("Getting cell data.");
	  //if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
	  //{ 
		//  cell.setCellType(XSSFCell.CELL_TYPE_STRING);
	  //}
	  //String cellData = cell.getStringCellValue();
	 // return cellData;
	  
	  
	  DataFormatter formatter=new DataFormatter();
	  
	  HSSFSheet excelSheet = excelWorkbook.getSheet(sheetName);
	  Cell = excelSheet.getRow(rowNumb).getCell(colNumb);
	  
	  // String CellData = Cell.getStringCellValue(); // return CellData;
	  
		if (Cell.getCellType() == CellType.STRING) {
			String cellData = formatter.formatCellValue(Cell);
			//System.out.println("getCellData :" + cellData);
			return cellData;
		} else if (Cell.getCellType() == CellType.NUMERIC) {
			FormulaEvaluator evaluator = excelWorkbook.getCreationHelper().createFormulaEvaluator();
			String cellData = formatter.formatCellValue(Cell, evaluator);
			//System.out.println("getCellData :" + cellData);
			return cellData;
		} else if (Cell.getCellType() == CellType.FORMULA) {
			FormulaEvaluator evaluator = excelWorkbook.getCreationHelper().createFormulaEvaluator();
			String cellData = formatter.formatCellValue(Cell, evaluator);
			//System.out.println("getCellData :" + cellData);
			return cellData;
		}

		else if (Cell.getCellType() == CellType.BLANK)

		{
			return "";
		}

		else {
			//System.out.println("getCellData :" + String.valueOf(Cell.getBooleanCellValue()));
			return String.valueOf(Cell.getBooleanCellValue());
		}

	} catch (Exception exp) {
		return "";
	}
}

//New Added to Get Data from Columns
	public synchronized String getCellColData(String testCaseName, int colNumb, String sheetName) throws Exception{
		
		try{
			  
			  //XSSFSheet excelSheet = excelWorkbook.getSheet(sheetName); 
			  //XSSFCell cell = excelSheet.getRow(rowNumb).getCell(colNumb);
			  //log.debug("Getting cell data.");
			  //if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC)
			  //{ 
				//  cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			  //}
			  //String cellData = cell.getStringCellValue();
			 // return cellData;
			  
			  
			  DataFormatter formatter=new DataFormatter();
			  
			  HSSFSheet excelSheet = excelWorkbook.getSheet(sheetName);
			  
			 // HSSFRow Row = null; 
				    			 
				// Making the object of excel row 
				int iRowNumCount = getNumberOfRows(sheetName); 
				int iExpectedTestCaseRowNum=0;
				/// ************* Testing
				for(int iRowNum=1;iRowNum<iRowNumCount+1;iRowNum++)
		    	{
		    		Cell cell=excelSheet.getRow(iRowNum).getCell(0);
		    		String RowTestNameCellData=cell.getStringCellValue().toString();
		    		
		    		if(RowTestNameCellData.equalsIgnoreCase((testCaseName).trim()))
		    		{
		    			iExpectedTestCaseRowNum=(int)iRowNum;
		    			System.out.println("getCellColData()-TestCase Name "+testCaseName+" Found in DataSheet "+sheetName+" At Row Num "+iExpectedTestCaseRowNum+"  For getCellColData:"+iExpectedTestCaseRowNum);
		    			break;
		    			
		    		}
		    	}
			  
			  
			  Cell = excelSheet.getRow(iExpectedTestCaseRowNum).getCell(colNumb);
			  
			  // String CellData = Cell.getStringCellValue(); // return CellData;
			  
				if (Cell.getCellType() == CellType.STRING) {
					String cellData = formatter.formatCellValue(Cell);
					//System.out.println("getCellColData :" + cellData);
					return cellData;
				} else if (Cell.getCellType() == CellType.NUMERIC) {
					FormulaEvaluator evaluator = excelWorkbook.getCreationHelper().createFormulaEvaluator();
					String cellData = formatter.formatCellValue(Cell, evaluator);
					//System.out.println("getCellColData :" + cellData);
					return cellData;
				} else if (Cell.getCellType() == CellType.FORMULA) {
					FormulaEvaluator evaluator = excelWorkbook.getCreationHelper().createFormulaEvaluator();
					String cellData = formatter.formatCellValue(Cell, evaluator);
					//System.out.println("getCellColData :" + cellData);
					return cellData;
				}

				else if (Cell.getCellType() == CellType.BLANK)

				{
					return "";
				}

				else {
					//System.out.println("getCellColData :" + String.valueOf(Cell.getBooleanCellValue()));
					return String.valueOf(Cell.getBooleanCellValue());
				}

			} catch (Exception exp) {
				return "";
			}
	}

	public static synchronized void clearColumnData(String sheetName, int colNumb, String excelFilePath) {
		int rowCount = getNumberOfRows(sheetName);
		HSSFRow row;
		HSSFSheet excelSheet = excelWorkbook.getSheet(sheetName);
		for (int i = 1; i < rowCount; i++) {
			HSSFCell cell = excelSheet.getRow(i).getCell(colNumb);
			if (cell == null) {
				row = excelSheet.getRow(i);
				cell = row.createCell(colNumb);
			}
			cell.setCellValue("");
		}
		log.debug("Clearing column " + colNumb + " of Sheet: " + sheetName);
		writingDataIntoFile(excelFilePath);
	}

	public synchronized void setCellData(String result, int rowNumb, int colNumb, String excelFilePath,
			String sheetName) {
		HSSFSheet excelSheet = excelWorkbook.getSheet(sheetName);
		HSSFRow row = excelSheet.getRow(rowNumb);
		HSSFCell cell = row.getCell(colNumb);
		log.debug("Setting results into the excel sheet.");
		if (cell == null) {
			cell = row.createCell(colNumb);
		}
		cell.setCellValue(result);
		log.debug("Setting value into cell[" + rowNumb + "][" + colNumb + "]: " + result);
		writingDataIntoFile(excelFilePath);
	}

	private static synchronized void writingDataIntoFile(String excelFilePath) {
		try {
			FileOutputStream fileOut = new FileOutputStream(excelFilePath);
			excelWorkbook.write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (Exception exp) {
			log.error("Exception occured in setCellData: ", exp);
		}
	}

public synchronized Map<Integer, String> getTestCasesToRun(String sheetName, int runModeColumn,int testCaseColumn) {
		Map<Integer, String> testListMap = new HashMap<Integer, String>();
		System.out.println(
				"sheetName runModeColumn testCaseColumn :" + sheetName + "_" + runModeColumn + "_" + testCaseColumn);
		try {
			int rowCount = getNumberOfRows(sheetName);
			System.out.println("Excel rowCount :" + rowCount);
			String testCase;
			for (int i = 1; i < rowCount; i++) {
				testCase = getTestCaseToRun(i, runModeColumn, testCaseColumn, sheetName);
				if (testCase != null) {
					testListMap.put(i, testCase);
				}
			}
		} catch (Exception e) {
			log.error("Exeception Occured while adding data to List:\n", e);
		}
		return testListMap;
	}

	private synchronized String getTestCaseToRun(int row, int runModeColumn, int testCaseColumn, String sheetName) {
		String testCaseName = null;
		try {
			if (getCellData(row, runModeColumn, sheetName).equalsIgnoreCase(RUN_MODE_YES)) {
				testCaseName = getCellData(row, testCaseColumn, sheetName).trim();
				log.debug("Test Case to Run: " + testCaseName);
			}
		} catch (Exception exp) {
			log.error("Exception occured in getTestCaseRow: ", exp);
		}
		return testCaseName;
	}
	
// New Added To get Sheet Row Entire Data Based on TestCaseName in Hash Map
			
	public synchronized List<Map<String,String>> getMapTestData(String sheetName,String testCaseID) throws IOException
	  {
		List<Map<String,String>>testDataAllRows=null;
		
		Map<String,String> testData=null;
		
		try 
		{
			
			//workbook = new HSSFWorkbook(excelFile);
			//HSSFSheet sheet = workbook.getSheet(sheetName);
			
			HSSFSheet sheet = excelWorkbook.getSheet(sheetName);
			  
			int lastRowNumber=sheet.getLastRowNum()+1;
			int lastColNumber=sheet.getRow(0).getLastCellNum();
			
		    // test Main For loop for List Based on Test Case Name
			int TCEndCounterRow=0;
			int startpoint=0;
			
			for( int count=1;count<lastRowNumber;count++)
			{
				//HSSFSheet excelSheet = workbook.getSheet(sheetName);
				
				Cell cellTest=sheet.getRow(count).getCell(0);
				String RowTestNameCellData=cellTest.getStringCellValue().toString();
				if(testCaseID.equalsIgnoreCase(RowTestNameCellData))
				   {
				    //
					System.out.println("TestCase "+testCaseID+" in DataSheet "+sheetName+" Found At Row :"+count);
						
					startpoint=startpoint+1;
					System.out.println("startpoint :"+startpoint);
					TCEndCounterRow=count;
					//
			      } // main if end
			} // Main For loop end	

	  	
	  	int TCstartPointer=(TCEndCounterRow-startpoint)+1;//plus 1
	  	int TCEndPointer=(TCstartPointer+startpoint)-1;//minus 1
	  	DataFormatter formatter=new DataFormatter();
	  	
	  	System.out.println("startpointStart :"+TCstartPointer);
	  	System.out.println("finalcount :"+TCEndPointer);
	  	
			//List<String> list=new ArrayList<String>();
			List list=new ArrayList();
			for(int i=0;i<lastColNumber;i++)
			{
				
				HSSFRow row=sheet.getRow(0);
				HSSFCell cell=row.getCell(i);
				// Formatting
				if(cell.getCellType()==CellType.STRING)
		       	{
					String rowHeader=formatter.formatCellValue(cell);
					System.out.println("rowHeader :"+rowHeader);
		       		list.add(rowHeader);
		       		
		       	}
		       	else if(cell.getCellType()==CellType.NUMERIC)
		       	{
		       		FormulaEvaluator evaluator=excelWorkbook.getCreationHelper().createFormulaEvaluator();
		       		String rowHeader=formatter.formatCellValue(cell,evaluator);
		       		System.out.println("rowHeader :"+rowHeader);
		       		list.add(rowHeader);
		       		
		       	}
		       	else if(cell.getCellType()==CellType.FORMULA)
		       	{
		       		FormulaEvaluator evaluator=excelWorkbook.getCreationHelper().createFormulaEvaluator();
		       		String rowHeader=formatter.formatCellValue(cell,evaluator);
		       		System.out.println("rowHeader :"+rowHeader);
		       		list.add(rowHeader);
		       		
		       	}
		       	
		       	else if(cell.getCellType()==CellType.BOOLEAN)
		       	{
		       		
		       		String rowHeader= String.valueOf(cell.getBooleanCellValue());
		       		System.out.println("rowHeader :"+rowHeader);
		       		list.add(rowHeader);
		       	}
				else if(cell.getCellType()==CellType.BLANK)
		       		
		       	{
						String rowHeader="";
						System.out.println("rowHeader :"+rowHeader);
						list.add(rowHeader);
		       	}
				
				// Data Formatting End
				
				//String rowHeader=cell.getStringCellValue().trim();
				//System.out.println("rowHeader :"+rowHeader);
				//list.add(rowHeader);
				
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
					
					// Data Formatting
					if(cell.getCellType()==CellType.STRING)
			       	{
						String colValue=formatter.formatCellValue(cell);
						System.out.println("colValue :"+colValue);
			       		testData.put((String) list.get(k), colValue);
			       		
			       	}
			       	else if(cell.getCellType()==CellType.NUMERIC)
			       	{
			       		FormulaEvaluator evaluator=excelWorkbook.getCreationHelper().createFormulaEvaluator();
			       		String colValue=formatter.formatCellValue(cell,evaluator);
			       		System.out.println("colValue :"+colValue);
			       		testData.put((String) list.get(k), colValue);
			       		
			       	}
			       	else if(cell.getCellType()==CellType.FORMULA)
			       	{
			       		FormulaEvaluator evaluator=excelWorkbook.getCreationHelper().createFormulaEvaluator();
			       		String colValue=formatter.formatCellValue(cell,evaluator);
			       		System.out.println("colValue :"+colValue);
			       		testData.put((String) list.get(k), colValue);
			       		
			       	}
			       	
			       	else if(cell.getCellType()==CellType.BOOLEAN)
			       	{
			       		
			       		String colValue= String.valueOf(cell.getBooleanCellValue());
			       		System.out.println("colValue :"+colValue);
			       		testData.put((String) list.get(k), colValue);
			       	}
					else if(cell.getCellType()==CellType.BLANK)
			       		
			       	{
							String colValue="";
							testData.put((String) list.get(k), colValue);
			       	}
										
					// Data Formatting End
										
					//String colValue=cell.getStringCellValue().trim();
					//System.out.println("colValue :"+colValue);
					//testData.put((String) list.get(k), colValue);
					
				}
				testDataAllRows.add(testData);
			}
					
		}
		 catch (Exception exp) {
				log.error("Exception occured in getMapTestData: ", exp);
		}
		
		return testDataAllRows;
	 }
	
		
	
//	
	
	
	
	
	
	
}
